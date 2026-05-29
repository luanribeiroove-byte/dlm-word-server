#!/usr/bin/env python3
"""
gerar_relatorio.py v3 — Gera o Relatório ETE Jardim Cajazeiras.

Mudanças nesta versão:
- Seção 4 (Detalhamento das Atividades) e o trecho 4.x do sumário são gerados
  dinamicamente. O usuário define a quantidade de inspeções (3, 4, 5, ...) e
  a quantidade de fotos por inspeção (1, 2, 3, ...) — tudo é montado no
  documento preservando o estilo do espelho.
- Tabela de NCs continua dinâmica (1+ linhas).
- Fotos são redimensionadas automaticamente preservando proporção.
- Imagens novas são empacotadas no .docx com as relações XML corretas.

Uso:
    python gerar_relatorio.py <dados.json> <pasta_saida> [pasta_skill] [pasta_assets]
"""

import json
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from pathlib import Path

# Importar o builder da Seção 4
sys.path.insert(0, str(Path(__file__).parent))
from secao4_builder import construir_secao4, construir_sumario_4x, escape_xml


# ============================================================
# UTILS
# ============================================================

def slugify(s: str) -> str:
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("/", "_").replace(" ", "_").lower()


# ============================================================
# CONFORMIDADE CONAMA 430/2011
# ============================================================

def _to_float(v):
    """Tenta converter um valor (string ou número) pra float. Retorna None se falhar."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        # Remove unidades, %, espaços e troca vírgula por ponto
        s = v.strip().replace(",", ".").replace("%", "")
        # Remove tudo que não é número, ponto ou sinal de menos
        s = re.sub(r"[^0-9.\-]", "", s)
        if not s or s in (".", "-", "-."):
            return None
        try:
            return float(s)
        except (ValueError, TypeError):
            return None
    return None


def _conformidade_dbo(p: dict) -> str:
    """
    DBO conforme CONAMA 430/2011: tratado ≤ 120 mg/L OU eficiência ≥ 60%.
    Retorna 'CONFORME', 'NÃO CONFORME' ou '—' se sem dados.
    """
    tratado = _to_float(p.get("tratado"))
    efic = _to_float(p.get("eficiencia"))
    if tratado is None and efic is None:
        return "—"
    # Regra dupla: basta uma satisfazer
    if tratado is not None and tratado <= 120:
        return "CONFORME"
    if efic is not None and efic >= 60:
        return "CONFORME"
    return "NÃO CONFORME"


def _conformidade_ph(p: dict) -> str:
    """pH conforme CONAMA 430/2011: 5 ≤ tratado ≤ 9."""
    tratado = _to_float(p.get("tratado"))
    if tratado is None:
        return "—"
    if 5 <= tratado <= 9:
        return "CONFORME"
    return "NÃO CONFORME"


def _conformidade_temp(p: dict) -> str:
    """Temperatura conforme CONAMA 430/2011: tratado < 40°C."""
    tratado = _to_float(p.get("tratado"))
    if tratado is None:
        return "—"
    if tratado < 40:
        return "CONFORME"
    return "NÃO CONFORME"


# ============================================================
# MODELO DASA — 2 ETEs, 11 parâmetros, regras CONAMA 430 Art. 21
# ============================================================

def _conformidade_eficiencia(p: dict) -> str:
    """
    Mostra a eficiência como texto para parâmetros sem limite direto CONAMA 430.
    Ex: '99,3% remoção'. Se não conseguir calcular, retorna '—'.
    """
    if not p:
        return "—"
    efic = p.get("eficiencia", "")
    if efic and efic != "—":
        # Garantir que tem "%" no final
        efic_str = str(efic).strip()
        if not efic_str.endswith("%"):
            efic_str = efic_str + "%"
        return f"{efic_str} remoção"
    return "—"




def _parse_valor_dasa(s):
    """
    Converte string em (valor_float, tem_menor_que, tem_maior_que).

    Exemplos:
      "7,12"        -> (7.12, False, False)
      "< 4,3"       -> (4.3, True, False)
      "> 3,0 × 10³" -> (3000.0, False, True)
      "1.080"       -> (1080.0, False, False)
      "2,8 × 10²"   -> (280.0, False, False)
      "Ausência"    -> (None, False, False)
    """
    if not s or not isinstance(s, str):
        return (None, False, False)

    s = s.strip()

    if s.lower() in ('ausência', 'ausencia', 'ausente'):
        return (None, False, False)

    tem_menor = '<' in s
    tem_maior = '>' in s
    s = s.replace('<', '').replace('>', '').strip()

    # Notação científica brasileira: "2,8 × 10²"
    m = re.match(r'(\d+[,.]?\d*)\s*[×x]\s*10\s*([²³⁴⁵⁶⁷⁸⁹¹⁰]+)', s)
    if m:
        base = float(m.group(1).replace(',', '.'))
        exp_map = {'⁰': 0, '¹': 1, '²': 2, '³': 3, '⁴': 4, '⁵': 5,
                   '⁶': 6, '⁷': 7, '⁸': 8, '⁹': 9}
        exp_str = m.group(2)
        exp = int(''.join(str(exp_map.get(c, '')) for c in exp_str) or '0')
        return (base * (10 ** exp), tem_menor, tem_maior)

    # Número simples (formato pt-BR): ponto = milhar, vírgula = decimal
    s = s.replace('.', '').replace(',', '.')
    try:
        return (float(s), tem_menor, tem_maior)
    except ValueError:
        return (None, tem_menor, tem_maior)


def _fmt_pct_dasa(valor, com_maior=False):
    """Formata porcentagem em pt-BR: 84.4 -> '84,4%' ou '> 84,4%' """
    txt = f"{valor:.1f}".replace('.', ',') + '%'
    if com_maior:
        txt = '> ' + txt
    return txt


def _eficiencia_dasa(ent_val, sai_val):
    """Calcula eficiência de remoção em %"""
    if ent_val is None or sai_val is None or ent_val <= 0:
        return None
    eff = (ent_val - sai_val) / ent_val * 100
    if eff < 0:
        return 0
    if eff > 100:
        return 100
    return eff


def calcular_situacao_dasa(parametro_id, entrada, saida):
    """
    Calcula a Situação (texto) de um parâmetro da tabela DASA.

    Baseado na CONAMA 430/2011 Art. 21 (esgoto sanitário) e Art. 16 II.

    parametro_id: PH, TEMP, OG, DBO, DQO, NTOTAL, NAMONIACAL, PTOTAL, SSED, SST, COLIF
    entrada, saida: strings com os valores

    Retorna: string com o texto da situação (cor é controlada no template).
    """
    ent_val, ent_lt, ent_gt = _parse_valor_dasa(entrada)
    sai_val, sai_lt, sai_gt = _parse_valor_dasa(saida)

    if parametro_id == 'PH':
        if sai_val is None:
            return '—'
        return 'Conforme' if 5.0 <= sai_val <= 9.0 else 'Não conforme'

    if parametro_id == 'TEMP':
        if sai_val is None:
            return '—'
        return 'Conforme' if sai_val < 40 else 'Não conforme'

    # CONAMA 430 Art. 21: até 100 mg/L
    if parametro_id == 'OG':
        if sai_val is None:
            return '—'
        return 'Conforme' if sai_val <= 100 else 'Não conforme'

    # CONAMA 430 Art. 21: máximo 120 mg/L OU remoção ≥ 60%
    if parametro_id == 'DBO':
        if sai_val is None:
            return '—'
        eff = _eficiencia_dasa(ent_val, sai_val)
        conforme_por_limite = sai_val <= 120
        conforme_por_eficiencia = eff is not None and eff >= 60
        if conforme_por_limite or conforme_por_eficiencia:
            if eff is not None:
                pct = _fmt_pct_dasa(eff, com_maior=ent_lt)
                return f'Conforme ({pct})'
            return 'Conforme'
        return 'Não conforme'

    # CONAMA 430 Art. 16 II: 20 mg/L N
    if parametro_id == 'NAMONIACAL':
        if sai_val is None:
            return '—'
        return 'Conforme' if sai_val <= 20 else 'Não conforme'

    # CONAMA 430 Art. 21: até 1,0 mL/L
    if parametro_id == 'SSED':
        if sai_val is None:
            return '—'
        return 'Conforme' if sai_val <= 1.0 else 'Não conforme'

    # Sem limite direto: mostrar eficiência
    if parametro_id == 'DQO':
        eff = _eficiencia_dasa(ent_val, sai_val)
        if eff is None:
            return '—'
        pct = _fmt_pct_dasa(eff, com_maior=(ent_lt or sai_lt))
        return f'{pct} remoção'

    if parametro_id == 'NTOTAL':
        # "Baixo" se ambos com "<" e valores iguais (limite de quantificação)
        if ent_lt and sai_lt and ent_val is not None and sai_val is not None:
            if abs(ent_val - sai_val) < 0.01:
                return 'Baixo'
        eff = _eficiencia_dasa(ent_val, sai_val)
        if eff is None or eff < 1:
            return '—'
        pct = _fmt_pct_dasa(eff, com_maior=sai_lt)
        return f'Redução {pct}'

    if parametro_id == 'PTOTAL':
        eff = _eficiencia_dasa(ent_val, sai_val)
        if eff is None:
            return '—'
        pct = _fmt_pct_dasa(eff, com_maior=sai_lt)
        return f'Redução {pct}'

    if parametro_id == 'SST':
        eff = _eficiencia_dasa(ent_val, sai_val)
        if eff is None:
            return '—'
        pct = _fmt_pct_dasa(eff, com_maior=sai_lt)
        return f'Redução {pct}'

    if parametro_id == 'COLIF':
        saida_norm = (saida or '').strip().lower()
        if saida_norm in ('ausência', 'ausencia', 'ausente'):
            return 'Excelente'
        if ent_val is not None and sai_val is not None:
            if ent_val > 1000 and sai_val > 1000:
                return 'Monitorar'
        return '—'

    return '—'


# IDs dos 11 parâmetros DASA (ordem da tabela)
PARAMETROS_DASA = [
    'PH', 'TEMP', 'OG', 'DBO', 'DQO',
    'NTOTAL', 'NAMONIACAL', 'PTOTAL', 'SSED', 'SST', 'COLIF'
]


def construir_mapa_dasa(dados: dict) -> dict:
    """
    Monta o dicionário de placeholders para o template DASA.

    Espera receber em `dados`:
      - cliente, empresa, cabecalho, introducao, quadro, inspecoes, ncs,
        consideracoes_paragrafos (igual ao modelo Simplificado)
      - ete1: { data_coleta, laudo_bruto, laudo_tratado, data_laudos,
                analise_tecnica, parametros: [{id, entrada, saida}, ...] }
      - ete2: { data_coleta, mes_laudo, laudo_bruto, laudo_tratado, data_laudos,
                analise_tecnica, observacao_tecnica, parametros: [...] }
    """
    m = {}

    # ===== Cliente / Empresa / Cabeçalho (mesma lógica do Simplificado) =====
    cliente = dados.get("cliente", {})
    nome_cliente = cliente.get("razao_social") or cliente.get("nome_exibicao") or "[A PREENCHER]"
    m["CLIENTE_NOME"] = nome_cliente
    m["CLIENTE_NOME_UPPER"] = nome_cliente.upper()
    m["CLIENTE_NOME_CABECALHO"] = cliente.get("nome_exibicao") or nome_cliente
    m["CLIENTE_CNPJ"] = cliente.get("cnpj", "[A PREENCHER]")
    m["CLIENTE_ENDERECO"] = cliente.get("endereco_completo") or cliente.get("endereco") or "[A PREENCHER]"
    m["TIPO_SISTEMA"] = cliente.get("tipo_sistema") or "Lodos Ativados — 2 unidades independentes (ETE 01 e ETE 02)"

    empresa = dados.get("empresa", {})
    m["EMPRESA_NOME"] = empresa.get("nome") or "DLM Saneamento e Engenharia Ltda."
    m["EMPRESA_NOME_UPPER"] = m["EMPRESA_NOME"].upper()
    m["RT_NOME"] = empresa.get("rt_nome") or "Diego Lopes Marinho"

    cab = dados.get("cabecalho", {})
    m["PERIODO"] = cab.get("periodo", "[A PREENCHER]")
    m["DATA_EMISSAO"] = cab.get("data_emissao", "[A PREENCHER]")

    intro = dados.get("introducao", {})
    m["INTRO_PERIODO"] = intro.get("periodo_intervencoes", "[A PREENCHER]")
    m["INTRO_ATIVIDADES"] = intro.get("atividades_narrativa", "[A PREENCHER]")
    m["LABORATORIO"] = dados.get("laboratorio") or cliente.get("laboratorio") or "[laboratório]"

    # ===== Quadro Resumo (Seção 3) =====
    inspecoes_lista = dados.get("inspecoes", [])
    quadro_default = dados.get("quadro", {})
    mes_ano_default = quadro_default.get("mes_ano", "[A PREENCHER]")
    classif_default = quadro_default.get("classificacao", "[A PREENCHER]")
    for i, insp in enumerate(inspecoes_lista, start=1):
        m[f"QUADRO{i}_MES_ANO"] = insp.get("mes_ano", mes_ano_default)
        m[f"QUADRO{i}_ATIVIDADE"] = insp.get("titulo", "[A PREENCHER]").rstrip(".").strip()
        m[f"QUADRO{i}_CLASSIFICACAO"] = insp.get("classificacao", classif_default)

    # ===== ETE 01 =====
    ete1 = dados.get("ete1", {})
    m["ETE1_DATA_COLETA"] = ete1.get("data_coleta", "[A PREENCHER]")
    m["ETE1_LAUDO_BRUTO"] = ete1.get("laudo_bruto", "[A PREENCHER]")
    m["ETE1_LAUDO_TRATADO"] = ete1.get("laudo_tratado", "[A PREENCHER]")
    m["ETE1_DATA_LAUDOS"] = ete1.get("data_laudos", "[A PREENCHER]")
    m["ETE1_ANALISE_TECNICA"] = ete1.get("analise_tecnica", "[A PREENCHER]")

    params_ete1 = {p.get("id"): p for p in ete1.get("parametros", [])}
    for pid in PARAMETROS_DASA:
        p = params_ete1.get(pid, {})
        ent = p.get("entrada", "—")
        sai = p.get("saida", "—")
        m[f"ETE1_{pid}_ENTRADA"] = ent
        m[f"ETE1_{pid}_SAIDA"] = sai
        m[f"ETE1_{pid}_SITUACAO"] = calcular_situacao_dasa(pid, ent, sai)

    # ===== ETE 02 =====
    ete2 = dados.get("ete2", {})
    m["ETE2_DATA_COLETA"] = ete2.get("data_coleta", "[A PREENCHER]")
    m["ETE2_MES_LAUDO"] = ete2.get("mes_laudo", "[A PREENCHER]")
    m["ETE2_LAUDO_BRUTO"] = ete2.get("laudo_bruto", "[A PREENCHER]")
    m["ETE2_LAUDO_TRATADO"] = ete2.get("laudo_tratado", "[A PREENCHER]")
    m["ETE2_DATA_LAUDOS"] = ete2.get("data_laudos", "[A PREENCHER]")
    m["ETE2_ANALISE_TECNICA"] = ete2.get("analise_tecnica", "[A PREENCHER]")
    m["ETE2_OBSERVACAO_TECNICA"] = ete2.get("observacao_tecnica", "")

    params_ete2 = {p.get("id"): p for p in ete2.get("parametros", [])}
    for pid in PARAMETROS_DASA:
        p = params_ete2.get(pid, {})
        ent = p.get("entrada", "—")
        sai = p.get("saida", "—")
        m[f"ETE2_{pid}_ENTRADA"] = ent
        m[f"ETE2_{pid}_SAIDA"] = sai
        m[f"ETE2_{pid}_SITUACAO"] = calcular_situacao_dasa(pid, ent, sai)

    # ===== NCs (até 3 — pela estrutura do template DASA) =====
    ncs = dados.get("nao_conformidades", [])
    for i in range(1, 4):
        if i <= len(ncs):
            nc = ncs[i-1]
            m[f"NC{i}_NUM"] = f"{i:02d}"
            m[f"NC{i}_DESC"] = nc.get("descricao", "—")
            m[f"NC{i}_MEDIDA"] = nc.get("medida", "—")
            m[f"NC{i}_PRIORIDADE"] = nc.get("prioridade", "—")
            m[f"NC{i}_STATUS"] = nc.get("status", "—")
        else:
            m[f"NC{i}_NUM"] = ""
            m[f"NC{i}_DESC"] = ""
            m[f"NC{i}_MEDIDA"] = ""
            m[f"NC{i}_PRIORIDADE"] = ""
            m[f"NC{i}_STATUS"] = ""

    # ===== Considerações finais (até 5 parágrafos — pelo template DASA) =====
    paragrafos = dados.get("consideracoes_paragrafos", [])
    for i in range(5):
        m[f"CONSID_P{i+1}"] = paragrafos[i] if i < len(paragrafos) else ""

    return m


def decidir_secoes_e_numerar(dados: dict) -> dict:
    """
    Decide quais seções aparecem no relatório e calcula a numeração dinâmica.

    Regras:
    - INTRODUÇÃO, DADOS, CONSIDERAÇÕES: sempre presentes
    - QUADRO e DETALHAMENTO: só se houver inspeções
    - NCS: só se houver não conformidades (vem ANTES de EFICIENCIA agora — Obs 3)
    - EFICIÊNCIA: só se houve campanha laboratorial
    - ANEXOS: só se houver campanha laboratorial com pelo menos um laudo informado
              (os anexos são os boletins/laudos do laboratório)

    Retorna: { 'presente': {nome: bool}, 'numero': {nome: int|None} }
    """
    inspecoes = dados.get("inspecoes", []) or []
    eficiencia = dados.get("eficiencia", {}) or {}
    ncs = dados.get("nao_conformidades", []) or []

    # ANEXOS = laudos da campanha. Se há campanha com laudo bruto OU tratado, há anexo.
    houve_campanha = bool(eficiencia.get("houve_campanha"))
    laudo_bruto = (eficiencia.get("laudo_bruto") or "").strip()
    laudo_tratado = (eficiencia.get("laudo_tratado") or "").strip()
    tem_anexos = houve_campanha and (bool(laudo_bruto) or bool(laudo_tratado))

    presente = {
        "INTRODUCAO": True,
        "DADOS": True,
        "QUADRO": len(inspecoes) > 0,
        "DETALHAMENTO": len(inspecoes) > 0,
        "NCS": len(ncs) > 0,                                    # antes de EFIC
        "EFICIENCIA": houve_campanha,
        "CONSIDERACOES": True,
        "ANEXOS": tem_anexos,
    }

    # Numeração: só seções presentes recebem número sequencial.
    # Ordem: INTRO, DADOS, QUADRO, DETALHAMENTO, NCS, EFICIENCIA, CONSIDERACOES, ANEXOS
    numero = {}
    contador = 0
    for nome in ["INTRODUCAO", "DADOS", "QUADRO", "DETALHAMENTO",
                 "NCS", "EFICIENCIA", "CONSIDERACOES", "ANEXOS"]:
        if presente[nome]:
            contador += 1
            numero[nome] = contador
        else:
            numero[nome] = None

    return {"presente": presente, "numero": numero}


def remover_secoes_ausentes(xml: str, decisao: dict) -> str:
    """
    Remove do XML as áreas marcadas das seções ausentes.
    Remove tanto a área do corpo (__SECAO_X_INI/FIM__) quanto a linha do sumário (__SUMLINE_X_INI/FIM__).
    """
    for nome in ["QUADRO", "DETALHAMENTO", "EFICIENCIA", "NCS", "ANEXOS"]:
        if not decisao["presente"].get(nome, False):
            xml = substituir_area_marcada(xml, f"__SECAO_{nome}_INI__", f"__SECAO_{nome}_FIM__", "")
            xml = substituir_area_marcada(xml, f"__SUMLINE_{nome}_INI__", f"__SUMLINE_{nome}_FIM__", "")

    # OBS 6: SECAO_LAB no parágrafo da Introdução é INLINE (runs dentro de um parágrafo).
    # Quando não há campanha (EFICIENCIA ausente), remove apenas os runs entre os
    # marcadores, preservando o restante do parágrafo da Introdução.
    if not decisao["presente"].get("EFICIENCIA", False):
        import re as _re
        # Remove tudo entre {{__SECAO_LAB_INI__}} e {{__SECAO_LAB_FIM__}} (inclusive)
        pattern = (
            r'<w:r>(?:(?!<w:r>).)*?\{\{__SECAO_LAB_INI__\}\}.*?'
            r'\{\{__SECAO_LAB_FIM__\}\}.*?</w:r>'
        )
        xml = _re.sub(pattern, '', xml, flags=_re.DOTALL)

    return xml


def limpar_marcadores_secao(xml: str) -> str:
    """
    Para seções PRESENTES, os marcadores __SECAO_X_INI/FIM__ e __SUMLINE_X_INI/FIM__ 
    ficaram no XML. Vamos limpar removendo os parágrafos que contêm esses marcadores.
    Trata também marcadores especiais (SECAO_LAB inline e SUMLINE_ANEXOS).
    """
    # Marcadores que ficam em PARÁGRAFOS próprios (envolvem áreas)
    paragrafo_markers = []
    for nome in ["QUADRO", "DETALHAMENTO", "EFICIENCIA", "NCS", "ANEXOS"]:
        for prefixo in ["SECAO", "SUMLINE"]:
            for sufixo in ["INI", "FIM"]:
                paragrafo_markers.append(f"__{prefixo}_{nome}_{sufixo}__")

    for marker in paragrafo_markers:
        while True:
            pos = xml.find("{{" + marker + "}}")
            if pos == -1:
                break
            p_ini = max(xml.rfind("<w:p>", 0, pos), xml.rfind("<w:p ", 0, pos))
            p_fim = xml.find("</w:p>", pos)
            if p_ini == -1 or p_fim == -1:
                xml = xml.replace("{{" + marker + "}}", "")
                break
            p_fim += len("</w:p>")
            xml = xml[:p_ini] + xml[p_fim:]

    # Marcadores INLINE (SECAO_LAB) — quando presente, remove só o run do marcador
    # (não remove o parágrafo todo, porque ele tem texto válido em volta)
    import re
    for marker in ["__SECAO_LAB_INI__", "__SECAO_LAB_FIM__"]:
        # Remove o <w:r>...{{marker}}...</w:r> inteiro
        pattern = r'<w:r>(?:(?!<w:r>).)*?\{\{' + re.escape(marker) + r'\}\}.*?</w:r>'
        xml = re.sub(pattern, '', xml, flags=re.DOTALL)

    return xml


def aplicar_substituicoes(xml: str, mapa: dict) -> str:
    for placeholder, valor in mapa.items():
        xml = xml.replace("{{" + placeholder + "}}", escape_xml(valor))
    return xml


# ============================================================
# AJUSTE DINÂMICO DA TABELA DE NCs
# ============================================================

def ajustar_linhas_nc(xml: str, num_ncs: int) -> str:
    """Adiciona/remove linhas da tabela de NCs conforme num_ncs (mesma lógica de antes)."""
    if num_ncs == 3:
        return xml

    linhas_originais = []
    for i in range(1, 4):
        marcador = "{{NC" + str(i) + "_DESC}}"
        pos = xml.find(marcador)
        if pos == -1:
            print(f"⚠ Marcador {marcador} não encontrado.")
            return xml
        inicio = xml.rfind("<w:tr ", 0, pos)
        if inicio == -1:
            inicio = xml.rfind("<w:tr>", 0, pos)
        fim = xml.find("</w:tr>", pos)
        if inicio == -1 or fim == -1:
            return xml
        fim += len("</w:tr>")
        linhas_originais.append((inicio, fim, xml[inicio:fim]))

    if num_ncs < 3:
        # Remove linhas excedentes (de trás pra frente)
        for i in range(2, num_ncs - 1, -1):
            inicio, fim, _ = linhas_originais[i]
            xml = xml[:inicio] + xml[fim:]
        return xml

    # num_ncs > 3: clonar a linha 3
    _, fim_linha3, xml_linha3 = linhas_originais[2]
    novas_linhas = []
    for idx in range(4, num_ncs + 1):
        clone = xml_linha3
        clone = clone.replace("{{NC3_NUM}}",        f"{{{{NC{idx}_NUM}}}}")
        clone = clone.replace("{{NC3_DESC}}",       f"{{{{NC{idx}_DESC}}}}")
        clone = clone.replace("{{NC3_MEDIDA}}",     f"{{{{NC{idx}_MEDIDA}}}}")
        clone = clone.replace("{{NC3_PRIORIDADE}}", f"{{{{NC{idx}_PRIORIDADE}}}}")
        clone = clone.replace("{{NC3_STATUS}}",     f"{{{{NC{idx}_STATUS}}}}")
        novas_linhas.append(clone)

    return xml[:fim_linha3] + "".join(novas_linhas) + xml[fim_linha3:]


# ============================================================
# AJUSTE DINÂMICO DA TABELA DO QUADRO RESUMO
# ============================================================

def remover_linhas_parametros_vazios(xml: str, mapa: dict) -> str:
    """
    Remove da tabela de eficiência as linhas de parâmetros sem valor.
    Regra (Opção A): a linha some apenas quando Bruto E Tratado estão vazios
    (em branco ou "—"). Deve rodar ANTES de aplicar_substituicoes, enquanto
    os placeholders {{XXX_BRUTO}} ainda existem como âncora.
    """
    def tem_numero(v):
        """True se o valor contém um número real (ignora >, <, espaços, vírgula decimal)."""
        if v is None:
            return False
        s = str(v).strip()
        if s in ("", "—", "-", "–", "N/D", "n/d", "N/A", "n/a"):
            return False
        # remove sinais de comparação e notação, troca vírgula por ponto
        s = s.replace("<", "").replace(">", "").replace("±", "").strip()
        s = s.replace(".", "").replace(",", ".")  # tolera 1.234,56 e 1,6E+7
        import re as _re
        return bool(_re.search(r"\d", s))

    # prefixo do placeholder de cada parâmetro da tabela
    prefixos = ["DBO", "DQO", "NTOTAL", "PTOTAL", "COLIFORMES", "PH", "TEMP"]

    for prefixo in prefixos:
        bruto = mapa.get(f"{prefixo}_BRUTO")
        tratado = mapa.get(f"{prefixo}_TRATADO")
        if tem_numero(bruto) or tem_numero(tratado):
            continue  # tem ao menos um valor numérico → mantém a linha

        # acha a linha <w:tr> que contém o placeholder _BRUTO desse parâmetro
        marcador = "{{" + prefixo + "_BRUTO}}"
        pos = xml.find(marcador)
        if pos == -1:
            continue
        inicio = xml.rfind("<w:tr ", 0, pos)
        if inicio == -1:
            inicio = xml.rfind("<w:tr>", 0, pos)
        fim = xml.find("</w:tr>", pos)
        if inicio == -1 or fim == -1:
            continue
        fim += len("</w:tr>")
        xml = xml[:inicio] + xml[fim:]

    return xml


def ajustar_linhas_quadro(xml: str, num_atividades: int) -> str:
    """
    Adiciona/remove linhas da tabela do Quadro Resumo conforme num_atividades.
    A linha base no template usa placeholders {{QUADRO1_*}}; cada linha extra
    é clonada e renumerada para QUADRO2_*, QUADRO3_*, etc.
    """
    if num_atividades == 1:
        return xml

    # Localizar a linha QUADRO1
    marcador = "{{QUADRO1_MES_ANO}}"
    pos = xml.find(marcador)
    if pos == -1:
        print(f"⚠ Marcador {marcador} não encontrado.")
        return xml
    inicio = xml.rfind("<w:tr ", 0, pos)
    if inicio == -1:
        inicio = xml.rfind("<w:tr>", 0, pos)
    fim = xml.find("</w:tr>", pos)
    if inicio == -1 or fim == -1:
        return xml
    fim += len("</w:tr>")
    xml_linha1 = xml[inicio:fim]

    if num_atividades == 0:
        # Caso degenerado: deixa a linha QUADRO1 (vai ser preenchida com placeholders sem valor)
        return xml

    # Clonar para idx 2..N
    novas_linhas = []
    for idx in range(2, num_atividades + 1):
        clone = xml_linha1
        clone = clone.replace("{{QUADRO1_MES_ANO}}",       f"{{{{QUADRO{idx}_MES_ANO}}}}")
        clone = clone.replace("{{QUADRO1_ATIVIDADE}}",     f"{{{{QUADRO{idx}_ATIVIDADE}}}}")
        clone = clone.replace("{{QUADRO1_CLASSIFICACAO}}", f"{{{{QUADRO{idx}_CLASSIFICACAO}}}}")
        novas_linhas.append(clone)

    return xml[:fim] + "".join(novas_linhas) + xml[fim:]


# ============================================================
# REGISTRO DE IMAGENS NOVAS NO PACOTE .DOCX
# ============================================================

def reservar_rids_para_fotos(unpacked: Path, qtd: int) -> list:
    """
    Reserva qtd novos rIds em document.xml.rels (sem conflitar com os existentes)
    e retorna a lista de rIds. Nota: o caller ainda precisa adicionar os
    Relationships reais com Target após copiar os arquivos.
    """
    rels_path = unpacked / "word" / "_rels" / "document.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")

    # Encontrar o maior rId atual
    ids_existentes = re.findall(r'Id="rId(\d+)"', rels_xml)
    max_id = max((int(i) for i in ids_existentes), default=0)

    novos = [f"rId{max_id + 1 + i}" for i in range(qtd)]
    return novos


def adicionar_relationships_imagens(unpacked: Path, mapa: list):
    """
    mapa = [(rid, nome_arquivo_destino), ...]
    Adiciona <Relationship> em document.xml.rels para cada imagem nova.
    """
    if not mapa:
        return

    rels_path = unpacked / "word" / "_rels" / "document.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")

    novos_rels = []
    for rid, nome in mapa:
        novos_rels.append(
            f'  <Relationship Id="{rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="media/{nome}"/>'
        )

    # Inserir antes do </Relationships>
    rels_xml = rels_xml.replace(
        "</Relationships>",
        "\n".join(novos_rels) + "\n</Relationships>"
    )
    rels_path.write_text(rels_xml, encoding="utf-8")


def copiar_fotos_para_media(unpacked: Path, mapa: list):
    """mapa = [(arquivo_origem, nome_destino, rid), ...]"""
    pasta_media = unpacked / "word" / "media"
    pasta_media.mkdir(exist_ok=True)
    for origem, destino, _ in mapa:
        shutil.copy(origem, pasta_media / destino)


# ============================================================
# CONSTRUÇÃO DO MAPA DE PLACEHOLDERS
# ============================================================

def atividades_para_narrativa(texto: str) -> str:
    """
    Transforma a string de atividades do quadro resumo em texto narrativo
    para encaixar na frase do P2 da introdução.

    Exemplos:
        "Auditoria da estrutura metálica; Verificação da bomba."
        → "auditoria da estrutura metálica e verificação da bomba"

        "Inspeção do painel; troca de relé; ETE inspecionada."
        → "inspeção do painel, troca de relé e ETE inspecionada"

    Regra:
        - Tira ponto final
        - Para cada item separado por ';': minúscula a primeira palavra,
          a menos que ela seja sigla (toda maiúscula com 2+ letras)
        - Junta com ',' e ' e ' no último
    """
    if not texto:
        return "[A PREENCHER]"

    t = texto.strip().rstrip(".").strip()
    if not t:
        return "[A PREENCHER]"

    def normalizar_item(item: str) -> str:
        item = item.strip()
        if not item:
            return item
        # Primeira palavra (até primeiro espaço)
        primeira_pal, _, resto = item.partition(" ")
        # É sigla? (2+ chars, todas alfa maiúsculas)
        if len(primeira_pal) >= 2 and primeira_pal.isalpha() and primeira_pal.isupper():
            return item  # preserva sigla
        # Caso normal: primeira letra minúscula
        return primeira_pal[0].lower() + primeira_pal[1:] + (" " + resto if resto else "")

    partes = [normalizar_item(p) for p in t.split(";") if p.strip()]
    if len(partes) <= 1:
        return partes[0] if partes else t
    if len(partes) == 2:
        return f"{partes[0]} e {partes[1]}"
    return ", ".join(partes[:-1]) + f" e {partes[-1]}"


def construir_mapa(dados: dict) -> dict:
    m = {}

    # ===== Numeração dinâmica das seções =====
    decisao = decidir_secoes_e_numerar(dados)
    for nome, num in decisao["numero"].items():
        m[f"NUM_{nome}"] = str(num) if num is not None else ""

    # ===== Dados do cliente =====
    cliente = dados.get("cliente", {})
    nome_cliente = cliente.get("razao_social") or cliente.get("nome") or "[A PREENCHER]"
    m["CLIENTE_NOME"] = nome_cliente
    m["CLIENTE_NOME_UPPER"] = nome_cliente.upper()
    m["CLIENTE_NOME_CABECALHO"] = cliente.get("nome_exibicao") or nome_cliente
    m["CLIENTE_CNPJ"] = cliente.get("cnpj", "[A PREENCHER]")
    m["CLIENTE_ENDERECO"] = cliente.get("endereco_completo") or cliente.get("endereco") or "[A PREENCHER]"
    m["TIPO_SISTEMA"] = cliente.get("tipo_sistema") or "ETE com Lodos Ativados e Aeração Prolongada"
    m["LABORATORIO"] = cliente.get("laboratorio") or "Bioagri Laboratórios Ltda. (Mérieux NutriSciences)"

    # ===== Dados da DLM (fixos) =====
    empresa = dados.get("empresa", {})
    m["EMPRESA_NOME"] = empresa.get("nome") or "DLM Saneamento e Engenharia Ltda."
    m["EMPRESA_NOME_UPPER"] = m["EMPRESA_NOME"].upper()
    m["EMPRESA_CNPJ"] = empresa.get("cnpj") or "29.745.355/0001-91"
    m["RT_NOME"] = empresa.get("rt_nome") or "Diego Lopes Marinho"

    # ===== Cabeçalho do relatório =====
    cab = dados.get("cabecalho", {})
    m["PERIODO"] = cab.get("periodo", "[A PREENCHER]")
    m["DATA_EMISSAO"] = cab.get("data_emissao", "[A PREENCHER]")

    intro = dados.get("introducao", {})
    m["INTRO_PERIODO"] = intro.get("periodo_intervencoes", "[A PREENCHER]")
    # INTRO_MESES_LAB: usa o explícito ou cai pro período do cabeçalho/introdução
    m["INTRO_MESES_LAB"] = (
        intro.get("meses_analises")
        or intro.get("periodo_intervencoes")
        or cab.get("periodo", "—")
    )

    # INTRO_ATIVIDADES: deriva dos títulos das inspeções, mas pode ser
    # sobrescrito por intro.atividades_narrativa se fornecido.
    inspecoes_lista = dados.get("inspecoes", [])
    titulos = [i.get("titulo", "") for i in inspecoes_lista if i.get("titulo")]
    intro_ativ_override = intro.get("atividades_narrativa")
    if intro_ativ_override:
        m["INTRO_ATIVIDADES"] = intro_ativ_override
    elif titulos:
        m["INTRO_ATIVIDADES"] = atividades_para_narrativa("; ".join(titulos))
    else:
        m["INTRO_ATIVIDADES"] = "[A PREENCHER]"

    m["PERIODO_ATIVIDADES"] = (
        dados.get("dados", {}).get("periodo_atividades")
        or cab.get("periodo", "—")
    )

    # Quadro Resumo (Seção 3): uma linha por inspeção.
    # Cada inspeção pode ter mes_ano e classificacao próprios; se omitidos,
    # usa fallbacks do bloco "quadro" (compatibilidade) ou "[A PREENCHER]".
    quadro_default = dados.get("quadro", {})
    mes_ano_default = quadro_default.get("mes_ano", "[A PREENCHER]")
    classif_default = quadro_default.get("classificacao", "[A PREENCHER]")

    for i, insp in enumerate(inspecoes_lista, start=1):
        titulo = insp.get("titulo", "[A PREENCHER]")
        # Tira ponto final se houver, pra não duplicar
        titulo_quadro = titulo.rstrip(".").strip()
        m[f"QUADRO{i}_MES_ANO"]       = insp.get("mes_ano", mes_ano_default)
        m[f"QUADRO{i}_ATIVIDADE"]     = titulo_quadro
        m[f"QUADRO{i}_CLASSIFICACAO"] = insp.get("classificacao", classif_default)

    e = dados.get("eficiencia", {})
    if e.get("houve_campanha"):
        m["EFIC_MES"] = e.get("mes_campanha", "[A PREENCHER]")
        m["DATA_AMOSTRAGEM"] = e.get("data_amostragem", "—")
        m["LAUDO_BRUTO"] = e.get("laudo_bruto", "—")
        m["LAUDO_TRATADO"] = e.get("laudo_tratado", "—")
        m["DATA_LAUDOS"] = e.get("data_laudos", "—")

        # NOVO: nome do laboratório (genérico, vem da extração da IA)
        m["LABORATORIO"] = e.get("laboratorio", "[laboratório acreditado]")

        params_map = {p["nome"]: p for p in e.get("parametros", [])}

        # Coliformes: aceita "Totais" OU "Termotolerantes" na mesma linha.
        # Rótulo da linha (COLIFORMES_LABEL) reflete o que veio do laudo.
        colif_termo = params_map.get("Coliformes Termotolerantes")
        colif_totais = params_map.get("Coliformes Totais")
        colif = colif_termo or colif_totais or {}
        if colif_termo:
            m["COLIFORMES_LABEL"] = "Coliformes Termotolerantes"
        elif colif_totais:
            m["COLIFORMES_LABEL"] = "Coliformes Totais"
        else:
            m["COLIFORMES_LABEL"] = "Coliformes Totais"  # padrão quando não há dado

        nomes_parametros = []  # pra montar a LISTA_PARAMETROS
        for nome_doc, prefixo in [
            ("DBO", "DBO"), ("DQO", "DQO"),
            ("Nitrogênio Total", "NTOTAL"), ("Fósforo Total", "PTOTAL"),
            ("__COLIF__", "COLIFORMES"),
            ("pH", "PH"), ("Temperatura", "TEMP"),
        ]:
            # Coliformes usa o dict resolvido acima (Totais ou Termotolerantes)
            p = colif if nome_doc == "__COLIF__" else params_map.get(nome_doc, {})
            m[f"{prefixo}_BRUTO"] = p.get("bruto", "—")
            m[f"{prefixo}_TRATADO"] = p.get("tratado", "—")
            if prefixo not in ("PH", "TEMP"):
                m[f"{prefixo}_EFIC"] = p.get("eficiencia", "—")
            # Acumula nomes pra LISTA_PARAMETROS (só se foi analisado)
            if p:
                rotulo = m["COLIFORMES_LABEL"] if nome_doc == "__COLIF__" else nome_doc
                nomes_parametros.append(rotulo)

        # NOVO: monta lista de parâmetros pra texto introdutório
        # ex: "DBO, DQO, Nitrogênio Total, Fósforo Total e Coliformes Totais"
        if nomes_parametros:
            if len(nomes_parametros) == 1:
                m["LISTA_PARAMETROS"] = nomes_parametros[0]
            else:
                m["LISTA_PARAMETROS"] = ", ".join(nomes_parametros[:-1]) + " e " + nomes_parametros[-1]
        else:
            m["LISTA_PARAMETROS"] = "DBO, DQO, Nitrogênio Total, Fósforo Total e Coliformes Totais"

        # NOVO: metodologia SMWW só aparece se a IA capturou (campo opcional)
        if e.get("metodologia_smww"):
            m["METODOLOGIA_OPCIONAL"] = ", seguindo as metodologias do Standard Methods for the Examination of Water and Wastewater (SMWW)"
        else:
            m["METODOLOGIA_OPCIONAL"] = ""

        # NOVO: calcula conformidade pra cada parâmetro com limite CONAMA 430
        # DBO: CONFORME se tratado ≤ 120 mg/L OU eficiência ≥ 60%
        m["DBO_CONFORMIDADE"] = _conformidade_dbo(params_map.get("DBO", {}))
        # pH: CONFORME se 5 ≤ tratado ≤ 9
        m["PH_CONFORMIDADE"] = _conformidade_ph(params_map.get("pH", {}))
        # Temperatura: CONFORME se tratado < 40
        m["TEMP_CONFORMIDADE"] = _conformidade_temp(params_map.get("Temperatura", {}))
        # Sem limite direto CONAMA: mostra eficiência (Obs 10)
        m["DQO_CONFORMIDADE"] = _conformidade_eficiencia(params_map.get("DQO", {}))
        m["NTOTAL_CONFORMIDADE"] = _conformidade_eficiencia(params_map.get("Nitrogênio Total", {}))
        m["PTOTAL_CONFORMIDADE"] = _conformidade_eficiencia(params_map.get("Fósforo Total", {}))
        m["COLIFORMES_CONFORMIDADE"] = _conformidade_eficiencia(colif)

        analises = e.get("analise_paragrafos", [])
        m["ANALISE_EFIC_1"] = analises[0] if len(analises) > 0 else ""
        m["ANALISE_EFIC_2"] = analises[1] if len(analises) > 1 else ""
        m["ANALISE_EFIC_3"] = analises[2] if len(analises) > 2 else ""
        # Placeholder de limpeza (texto fixo do template que não usamos mais)
        m["__SECAO5_LIMPAR_1__"] = ""
    else:
        m["EFIC_MES"] = "—"
        m["DATA_AMOSTRAGEM"] = "—"
        m["LAUDO_BRUTO"] = "—"
        m["LAUDO_TRATADO"] = "—"
        m["DATA_LAUDOS"] = "—"
        m["LABORATORIO"] = "—"
        m["LISTA_PARAMETROS"] = "—"
        m["METODOLOGIA_OPCIONAL"] = ""
        for prefixo in ["DBO", "DQO", "NTOTAL", "PTOTAL", "COLIFORMES", "PH", "TEMP"]:
            m[f"{prefixo}_BRUTO"] = "—"
            m[f"{prefixo}_TRATADO"] = "—"
            if prefixo not in ("PH", "TEMP"):
                m[f"{prefixo}_EFIC"] = "—"
        # Conformidades vazias quando não há campanha
        m["DBO_CONFORMIDADE"] = "—"
        m["PH_CONFORMIDADE"] = "—"
        m["TEMP_CONFORMIDADE"] = "—"
        m["DQO_CONFORMIDADE"] = "—"
        m["NTOTAL_CONFORMIDADE"] = "—"
        m["PTOTAL_CONFORMIDADE"] = "—"
        m["COLIFORMES_CONFORMIDADE"] = "—"
        m["COLIFORMES_LABEL"] = "Coliformes Totais"
        m["ANALISE_EFIC_1"] = "Não foi realizada campanha de monitoramento neste período."
        m["ANALISE_EFIC_2"] = ""
        m["ANALISE_EFIC_3"] = ""
        m["__SECAO5_LIMPAR_1__"] = ""

    # NCs (suporte a quantidade variável)
    ncs = dados.get("nao_conformidades", [])
    for i, nc in enumerate(ncs, start=1):
        m[f"NC{i}_NUM"] = f"{i:02d}"
        m[f"NC{i}_DESC"] = nc.get("descricao", "—")
        m[f"NC{i}_MEDIDA"] = nc.get("medida", "—")
        m[f"NC{i}_PRIORIDADE"] = nc.get("prioridade", "—")
        m[f"NC{i}_STATUS"] = nc.get("status", "—")

    # ===== Anexos (Seção 8) =====
    # Os anexos são os boletins de análise da campanha laboratorial.
    # Usamos dois placeholders separados (cada um num parágrafo próprio do template)
    # pra evitar problemas com quebra de linha dentro de uma string única.
    # A seção inteira é controlada pelos marcadores __SECAO_ANEXOS_INI/FIM__
    # (decidir_secoes_e_numerar decide se aparece).
    laudo_b = (e.get("laudo_bruto") or "").strip() if e.get("houve_campanha") else ""
    laudo_t = (e.get("laudo_tratado") or "").strip() if e.get("houve_campanha") else ""
    m["ANEXO_1"] = f"• Boletim de Análise nº {laudo_b} — Efluente Bruto" if laudo_b else ""
    m["ANEXO_2"] = f"• Boletim de Análise nº {laudo_t} — Efluente Tratado" if laudo_t else ""

    # Considerações finais (7 parágrafos)
    paragrafos = dados.get("consideracoes_paragrafos", [])
    for i in range(7):
        m[f"CONSID_P{i+1}"] = paragrafos[i] if i < len(paragrafos) else ""

    return m


# ============================================================
# SUBSTITUIÇÃO DAS ÁREAS DINÂMICAS
# ============================================================

# ============================================================
# COLORAÇÃO CONDICIONAL DE CÉLULAS (Prioridade + Conformidade)
# ============================================================

# Mapa: cor única no template → (placeholder lógico, função pra calcular cor real)
# Servidor substitui essas cores únicas pelas cores reais conforme o valor

CORES_PLACEHOLDER = {
    # Prioridade NCs
    "AAA001": "NC1_PRIORIDADE",
    "AAA002": "NC2_PRIORIDADE",
    "AAA003": "NC3_PRIORIDADE",
    # Conformidade Eficiência
    "AAA101": "DBO_CONFORMIDADE",
    "AAA102": "DQO_CONFORMIDADE",
    "AAA103": "NTOTAL_CONFORMIDADE",
    "AAA104": "PTOTAL_CONFORMIDADE",
    "AAA105": "COLIFORMES_CONFORMIDADE",
    "AAA106": "PH_CONFORMIDADE",
    "AAA107": "TEMP_CONFORMIDADE",
}

# Cores reais (hex sem #)
COR_VERMELHO = "F8D7DA"   # vermelho claro pra fundo (Urgente, Não Conforme)
COR_AMARELO  = "FFF3CD"   # amarelo claro (Alta prioridade)
COR_VERDE    = "D4EDDA"   # verde claro (Conforme)
COR_NEUTRA   = "FFFFFF"   # branco (sem destaque)


def _cor_prioridade(valor: str) -> str:
    """Retorna a cor hex baseada na prioridade da NC."""
    v = (valor or "").strip().lower()
    if v == "urgente":
        return COR_VERMELHO
    if v == "alta":
        return COR_AMARELO
    return COR_NEUTRA


def _cor_conformidade(valor: str) -> str:
    """Retorna a cor hex baseada no texto de conformidade.
    
    - 'CONFORME' (com ou sem %) → verde
    - 'NÃO CONFORME' → vermelho
    - qualquer outro (descritivo: '99% remoção', '—', etc) → sem cor
    """
    if not valor:
        return COR_NEUTRA
    v = valor.strip().lower()
    # "não conforme" tem que ser checado ANTES de "conforme"
    if "não conforme" in v or "nao conforme" in v:
        return COR_VERMELHO
    if "conforme" in v:
        return COR_VERDE
    return COR_NEUTRA


def aplicar_cores_celulas(xml: str, mapa: dict) -> str:
    """
    Substitui as cores únicas (AAA001, AAA002, ...) pelas cores reais 
    calculadas a partir dos valores em `mapa`.
    
    Deve ser chamado APÓS aplicar_substituicoes (que troca placeholders 
    de texto) — assim sabemos os valores finais de prioridade/conformidade.
    """
    for cor_unica, placeholder_logico in CORES_PLACEHOLDER.items():
        valor = mapa.get(placeholder_logico, "")
        if "PRIORIDADE" in placeholder_logico:
            cor_real = _cor_prioridade(valor)
        else:  # CONFORMIDADE
            cor_real = _cor_conformidade(valor)
        # Substituir a cor única pela cor real no shd
        xml = xml.replace(
            f'<w:shd w:val="clear" w:color="auto" w:fill="{cor_unica}"/>',
            f'<w:shd w:val="clear" w:color="auto" w:fill="{cor_real}"/>'
        )
    return xml




def substituir_area_marcada(xml: str, token_inicio: str, token_fim: str, novo_conteudo: str) -> str:
    """
    Os marcadores são parágrafos do tipo <w:p>...{{__TOKEN__}}...</w:p>.
    A função encontra os dois parágrafos marcadores e substitui tudo entre eles
    (inclusive os próprios parágrafos marcadores) pelo novo_conteudo.
    """
    # Encontrar o parágrafo que contém o token de início
    placeholder_inicio = "{{" + token_inicio + "}}"
    placeholder_fim = "{{" + token_fim + "}}"

    pos_token_ini = xml.find(placeholder_inicio)
    pos_token_fim = xml.find(placeholder_fim)
    if pos_token_ini == -1 or pos_token_fim == -1:
        print(f"⚠ Marcadores {token_inicio}/{token_fim} não encontrados no template.")
        return xml

    # Achar limites dos parágrafos
    # IMPORTANTE: buscar especificamente "<w:p>" ou "<w:p " (com espaço),
    # NÃO confundir com "<w:pPr>" (propriedades do parágrafo)
    inicio_p_ini = max(
        xml.rfind("<w:p>", 0, pos_token_ini),
        xml.rfind("<w:p ", 0, pos_token_ini)
    )
    fim_p_fim = xml.find("</w:p>", pos_token_fim) + len("</w:p>")

    return xml[:inicio_p_ini] + novo_conteudo + xml[fim_p_fim:]


# ============================================================
# CONVERSÃO PDF
# ============================================================

def converter_pdf(caminho_docx: Path, pasta_saida: Path):
    try:
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf",
             "--outdir", str(pasta_saida), str(caminho_docx)],
            check=True, capture_output=True, timeout=180,
        )
        return pasta_saida / (caminho_docx.stem + ".pdf")
    except Exception as e:
        print(f"⚠ Falha ao gerar PDF: {e}")
        return None


# ============================================================
# MAIN
# ============================================================

def main():
    if len(sys.argv) < 3:
        print("Uso: python gerar_relatorio.py <dados.json> <pasta_saida> [pasta_skill] [pasta_assets]")
        sys.exit(1)

    caminho_json = Path(sys.argv[1])
    pasta_saida = Path(sys.argv[2])
    pasta_skill = Path(sys.argv[3]) if len(sys.argv) > 3 else Path(__file__).parent.parent
    pasta_assets = Path(sys.argv[4]) if len(sys.argv) > 4 else caminho_json.parent

    pasta_saida.mkdir(parents=True, exist_ok=True)
    template = pasta_skill / "assets" / "template.docx"
    if not template.exists():
        print(f"❌ Template não encontrado: {template}")
        sys.exit(1)

    dados = json.loads(caminho_json.read_text(encoding="utf-8"))

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        unpacked = tmp / "unpacked"

        with zipfile.ZipFile(template, "r") as z:
            z.extractall(unpacked)

        doc_xml_path = unpacked / "word" / "document.xml"
        xml = doc_xml_path.read_text(encoding="utf-8")

        # 1. Construir Seção 4 dinâmica
        inspecoes = dados.get("inspecoes", [])
        if not inspecoes:
            print("⚠ Nenhuma inspeção fornecida. Seção 4 ficará vazia.")
            secao4_xml = ""
            sumario_4x_xml = ""
            imagens_a_processar = []
        else:
            # Contar fotos totais para reservar rIds
            total_fotos = sum(len(i.get("fotos", [])) for i in inspecoes)
            rids_disponiveis = reservar_rids_para_fotos(unpacked, total_fotos) if total_fotos else []

            secao4_xml, imagens_a_processar = construir_secao4(
                inspecoes, pasta_assets, unpacked, rids_disponiveis
            )
            sumario_4x_xml = construir_sumario_4x(inspecoes)

        # 2. Substituir áreas marcadas
        xml = substituir_area_marcada(xml, "__SUMARIO_4X_INICIO__", "__SUMARIO_4X_FIM__", sumario_4x_xml)
        xml = substituir_area_marcada(xml, "__SECAO4_INICIO__", "__SECAO4_FIM__", secao4_xml)

        # 2b. Remover seções ausentes e limpar marcadores das presentes
        decisao = decidir_secoes_e_numerar(dados)
        xml = remover_secoes_ausentes(xml, decisao)
        xml = limpar_marcadores_secao(xml)

        # 3a. Ajuste dinâmico da tabela do Quadro Resumo (1 linha por inspeção)
        num_atividades = len(dados.get("inspecoes", []))
        if num_atividades >= 1:
            xml = ajustar_linhas_quadro(xml, num_atividades)

        # 3b. Ajuste dinâmico da tabela de NCs
        num_ncs = len(dados.get("nao_conformidades", []))
        if num_ncs == 0:
            num_ncs = 3
        xml = ajustar_linhas_nc(xml, num_ncs)

        # 4. Substituir placeholders simples
        mapa = construir_mapa(dados)
        # 4b. Remover linhas de parâmetros sem valor (Bruto e Tratado vazios)
        xml = remover_linhas_parametros_vazios(xml, mapa)
        xml = aplicar_substituicoes(xml, mapa)
        doc_xml_path.write_text(xml, encoding="utf-8")

        # Aplicar substituições também no header e footer
        for nome_arq in ["header1.xml", "header2.xml", "footer1.xml", "footer2.xml"]:
            arq_path = unpacked / "word" / nome_arq
            if arq_path.exists():
                xml_aux = arq_path.read_text(encoding="utf-8")
                xml_aux = aplicar_substituicoes(xml_aux, mapa)
                arq_path.write_text(xml_aux, encoding="utf-8")

        # 5. Copiar fotos novas e adicionar Relationships
        if imagens_a_processar:
            copiar_fotos_para_media(unpacked, imagens_a_processar)
            mapa_rels = [(rid, nome) for _, nome, rid in imagens_a_processar]
            adicionar_relationships_imagens(unpacked, mapa_rels)
            print(f"✓ {len(imagens_a_processar)} fotos novas adicionadas")

        # 6. Repacotar
        sufixo = slugify(dados.get("cabecalho", {}).get("periodo", "mensal")) or "mensal"
        nome_base = f"Relatorio_ETE_Cajazeiras_{sufixo}"
        caminho_docx = pasta_saida / f"{nome_base}.docx"
        if caminho_docx.exists():
            caminho_docx.unlink()

        with zipfile.ZipFile(caminho_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for arquivo in unpacked.rglob("*"):
                if arquivo.is_file():
                    arcname = str(arquivo.relative_to(unpacked))
                    zout.write(arquivo, arcname)

        print(f"✓ DOCX: {caminho_docx}")

        caminho_pdf = converter_pdf(caminho_docx, pasta_saida)
        if caminho_pdf:
            print(f"✓ PDF:  {caminho_pdf}")

    print("\n✅ Relatório finalizado.")


if __name__ == "__main__":
    main()
