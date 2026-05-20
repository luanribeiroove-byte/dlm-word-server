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


def decidir_secoes_e_numerar(dados: dict) -> dict:
    """
    Decide quais seções aparecem no relatório e calcula a numeração dinâmica.

    Regras:
    - INTRODUÇÃO, DADOS, CONSIDERAÇÕES: sempre presentes
    - QUADRO e DETALHAMENTO: só se houver inspeções
    - EFICIÊNCIA: só se houve campanha laboratorial
    - NCS: só se houver não conformidades

    Retorna: { 'presente': {nome: bool}, 'numero': {nome: int|None} }
    """
    inspecoes = dados.get("inspecoes", []) or []
    eficiencia = dados.get("eficiencia", {}) or {}
    ncs = dados.get("nao_conformidades", []) or []

    presente = {
        "INTRODUCAO": True,
        "DADOS": True,
        "QUADRO": len(inspecoes) > 0,
        "DETALHAMENTO": len(inspecoes) > 0,
        "EFICIENCIA": bool(eficiencia.get("houve_campanha")),
        "NCS": len(ncs) > 0,
        "CONSIDERACOES": True,
    }

    # Calcula numeração: só seções presentes recebem número sequencial
    numero = {}
    contador = 0
    for nome in ["INTRODUCAO", "DADOS", "QUADRO", "DETALHAMENTO", "EFICIENCIA", "NCS", "CONSIDERACOES"]:
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
    for nome in ["QUADRO", "DETALHAMENTO", "EFICIENCIA", "NCS"]:
        if not decisao["presente"][nome]:
            xml = substituir_area_marcada(xml, f"__SECAO_{nome}_INI__", f"__SECAO_{nome}_FIM__", "")
            xml = substituir_area_marcada(xml, f"__SUMLINE_{nome}_INI__", f"__SUMLINE_{nome}_FIM__", "")
    return xml


def limpar_marcadores_secao(xml: str) -> str:
    """
    Para seções PRESENTES, os marcadores __SECAO_X_INI/FIM__ e __SUMLINE_X_INI/FIM__ 
    ficaram no XML. Vamos limpar removendo os parágrafos que contêm esses marcadores.
    """
    # Remove parágrafos que contêm apenas marcadores
    for nome in ["QUADRO", "DETALHAMENTO", "EFICIENCIA", "NCS"]:
        for prefixo in ["SECAO", "SUMLINE"]:
            for sufixo in ["INI", "FIM"]:
                marker = f"__{prefixo}_{nome}_{sufixo}__"
                # Procura o parágrafo que contém o marker e remove
                while True:
                    pos = xml.find("{{" + marker + "}}")
                    if pos == -1:
                        break
                    # Acha o <w:p> que contém
                    p_ini = max(xml.rfind("<w:p>", 0, pos), xml.rfind("<w:p ", 0, pos))
                    p_fim = xml.find("</w:p>", pos)
                    if p_ini == -1 or p_fim == -1:
                        # Não achou — substitui só o texto pra não travar
                        xml = xml.replace("{{" + marker + "}}", "")
                        break
                    p_fim += len("</w:p>")
                    xml = xml[:p_ini] + xml[p_fim:]
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
    m["INTRO_MESES_LAB"] = intro.get("meses_analises", "[A PREENCHER]")

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

    m["PERIODO_ATIVIDADES"] = dados.get("dados", {}).get("periodo_atividades", "[A PREENCHER]")

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
        m["DATA_AMOSTRAGEM"] = e.get("data_amostragem", "[A PREENCHER]")
        m["LAUDO_BRUTO"] = e.get("laudo_bruto", "[A PREENCHER]")
        m["LAUDO_TRATADO"] = e.get("laudo_tratado", "[A PREENCHER]")
        m["DATA_LAUDOS"] = e.get("data_laudos", "[A PREENCHER]")

        params_map = {p["nome"]: p for p in e.get("parametros", [])}
        for nome_doc, prefixo in [
            ("DBO", "DBO"), ("DQO", "DQO"),
            ("Nitrogênio Total", "NTOTAL"), ("Fósforo Total", "PTOTAL"),
            ("Coliformes Totais", "COLIFORMES"),
            ("pH", "PH"), ("Temperatura", "TEMP"),
        ]:
            p = params_map.get(nome_doc, {})
            m[f"{prefixo}_BRUTO"] = p.get("bruto", "—")
            m[f"{prefixo}_TRATADO"] = p.get("tratado", "—")
            if prefixo not in ("PH", "TEMP"):
                m[f"{prefixo}_EFIC"] = p.get("eficiencia", "—")

        analises = e.get("analise_paragrafos", [])
        m["ANALISE_EFIC_1"] = analises[0] if len(analises) > 0 else "[A PREENCHER]"
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
        for prefixo in ["DBO", "DQO", "NTOTAL", "PTOTAL", "COLIFORMES", "PH", "TEMP"]:
            m[f"{prefixo}_BRUTO"] = "—"
            m[f"{prefixo}_TRATADO"] = "—"
            if prefixo not in ("PH", "TEMP"):
                m[f"{prefixo}_EFIC"] = "—"
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

    # Considerações finais (7 parágrafos)
    paragrafos = dados.get("consideracoes_paragrafos", [])
    for i in range(7):
        m[f"CONSID_P{i+1}"] = paragrafos[i] if i < len(paragrafos) else ""

    return m


# ============================================================
# SUBSTITUIÇÃO DAS ÁREAS DINÂMICAS
# ============================================================

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
