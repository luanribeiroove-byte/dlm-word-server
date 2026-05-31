"""
app.py — Servidor Flask que gera relatórios DLM em .docx a partir de JSON.

Endpoints:
    GET  /            → Health check
    POST /gerar       → Recebe JSON, retorna .docx pra download

Como rodar localmente:
    pip install -r requirements.txt
    python app.py

Como rodar no Render:
    O Render usa o `Procfile` ou `Start Command` definido em runtime.
    Recomendado:  gunicorn app:app
"""

import json
import tempfile
import traceback
from pathlib import Path

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

# Importa o gerador (mesma pasta)
import sys
sys.path.insert(0, str(Path(__file__).parent))
from gerar_relatorio import (
    construir_mapa, aplicar_substituicoes,
    ajustar_linhas_nc, ajustar_linhas_quadro,
    reservar_rids_para_fotos, adicionar_relationships_imagens,
    copiar_fotos_para_media, substituir_area_marcada,
    decidir_secoes_e_numerar, remover_secoes_ausentes, limpar_marcadores_secao,
    construir_mapa_dasa,
    aplicar_cores_celulas,
)
from secao4_builder import construir_secao4, construir_sumario_4x
import zipfile

app = Flask(__name__)
CORS(app)  # Permite chamadas do app web em qualquer origem

ROOT = Path(__file__).parent
TEMPLATE = ROOT / "template_ete_simplificado.docx"
TEMPLATE_DASA = ROOT / "template_ete_dasa.docx"

# Mapa de modelos → template + fluxo de geração.
# Quando um modelo ainda não tem template próprio, usa fallback pro Simplificado.
TEMPLATES_DISPONIVEIS = {
    "intermediario_dasa": TEMPLATE_DASA,
    # Futuros: descomentar conforme forem criados
    # "industria_ortobom": ROOT / "template_ortobom.docx",
    # "condominio_eta_verana": ROOT / "template_verana.docx",
    # "hospitalar_menandro": ROOT / "template_menandro.docx",
    # "hospitalar_hcc": ROOT / "template_hcc.docx",
}


def escolher_template(modelo_padrao):
    """
    Retorna (caminho_template, modelo_efetivo).

    Se o modelo tem template próprio, usa esse template.
    Se não tem, FAZ FALLBACK pro Simplificado com aviso no log.
    Isso garante que o app sempre consegue gerar algum relatório,
    mesmo se um cliente foi marcado com modelo cujo template ainda
    não foi criado.
    """
    if modelo_padrao in TEMPLATES_DISPONIVEIS:
        template = TEMPLATES_DISPONIVEIS[modelo_padrao]
        if template.exists():
            return (template, modelo_padrao)
        print(f"⚠ Template {template.name} não encontrado, caindo no Simplificado")
    elif modelo_padrao and modelo_padrao != "simplificado":
        # Modelo válido no banco mas sem template ainda
        print(f"⚠ Modelo '{modelo_padrao}' sem template próprio, usando Simplificado")

    return (TEMPLATE, "simplificado")


@app.route("/", methods=["GET"])
def health():
    """Health check. O Render usa esse endpoint pra ver se o serviço tá vivo."""
    templates_status = {
        "simplificado": TEMPLATE.exists(),
    }
    for modelo, caminho in TEMPLATES_DISPONIVEIS.items():
        templates_status[modelo] = caminho.exists()

    return jsonify({
        "ok": True,
        "service": "dlm-word-generator",
        "version": "2.0.0",
        "templates": templates_status,
        "fallback": "Modelos sem template caem no Simplificado automaticamente"
    })


@app.route("/gerar", methods=["POST"])
def gerar_relatorio():
    """
    Recebe JSON com os dados e retorna o .docx gerado.

    Body esperado: JSON com a estrutura completa do relatório
    (ver dados_teste.json como exemplo).

    Retorna: arquivo .docx com Content-Disposition pra download.
    """
    try:
        dados = request.get_json(force=True)
        if not dados:
            return jsonify({"error": "Body vazio ou JSON inválido"}), 400

        # ===== Detectar modelo (vem do cliente) =====
        modelo_padrao_solicitado = (
            dados.get("modelo_padrao")
            or dados.get("cliente", {}).get("modelo_padrao")
            or "simplificado"
        )
        template_path, modelo_efetivo = escolher_template(modelo_padrao_solicitado)
        # modelo_efetivo é o que vai ser usado na geração (pode ser diferente
        # do solicitado se houve fallback pro Simplificado)
        modelo_padrao = modelo_efetivo

        # Nome do arquivo de saída
        cliente = dados.get("cabecalho", {}).get("cliente") or \
                  dados.get("cliente", {}).get("nome_exibicao", "Relatorio")
        periodo = dados.get("cabecalho", {}).get("periodo", "")
        nome_arquivo = f"Relatorio_{cliente.replace(' ', '_')}_{periodo.replace('/', '_')}.docx"
        nome_arquivo = "".join(c if c.isalnum() or c in "._-" else "_" for c in nome_arquivo)

        # Gerar em pasta temporária
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            unpacked = tmp / "unpacked"

            # Extrair template
            with zipfile.ZipFile(template_path, "r") as z:
                z.extractall(unpacked)

            doc_xml_path = unpacked / "word" / "document.xml"
            xml = doc_xml_path.read_text(encoding="utf-8")

            # ===== Roteamento por modelo =====
            if modelo_padrao == "intermediario_dasa":
                # DASA: substitui só os placeholders, sem Seção 4 dinâmica
                # nem ajuste de tabelas (o template DASA já tem estrutura fixa)
                mapa = construir_mapa_dasa(dados)
                xml = aplicar_substituicoes(xml, mapa)

                # Remover marcadores de seção 4 do template DASA (texto estático)
                texto_stub = ('Foram executadas, no período de {{INTRO_PERIODO}}, '
                              'atividades de manutenção preventiva e corretiva nas duas '
                              'estações de tratamento de esgoto da unidade DASA-Canela. '
                              'As intervenções foram organizadas conforme detalhamento a seguir.')
                # já foi substituído pelo INTRO_PERIODO no aplicar_substituicoes
                xml = xml.replace("{{__SECAO4_INICIO__}}", "").replace("{{__SECAO4_FIM__}}", "")

                doc_xml_path.write_text(xml, encoding="utf-8")

                # Header e footer
                for nome_arq in ["header1.xml", "header2.xml", "footer1.xml", "footer2.xml"]:
                    arq_path = unpacked / "word" / nome_arq
                    if arq_path.exists():
                        xml_aux = arq_path.read_text(encoding="utf-8")
                        xml_aux = aplicar_substituicoes(xml_aux, mapa)
                        arq_path.write_text(xml_aux, encoding="utf-8")

            else:
                # === Fluxo SIMPLIFICADO (legado, mantido sem alterações) ===
                inspecoes = dados.get("inspecoes", [])
                if not inspecoes:
                    secao4_xml = ""
                    sumario_4x_xml = ""
                    imagens_a_processar = []
                else:
                    total_fotos = sum(len(i.get("fotos", [])) for i in inspecoes)
                    rids_disponiveis = reservar_rids_para_fotos(unpacked, total_fotos) if total_fotos else []
                    secao4_xml, imagens_a_processar = construir_secao4(
                        inspecoes, tmp, unpacked, rids_disponiveis
                    )
                    sumario_4x_xml = construir_sumario_4x(inspecoes)

                xml = substituir_area_marcada(xml, "__SUMARIO_4X_INICIO__", "__SUMARIO_4X_FIM__", sumario_4x_xml)
                xml = substituir_area_marcada(xml, "__SECAO4_INICIO__", "__SECAO4_FIM__", secao4_xml)

                decisao = decidir_secoes_e_numerar(dados)
                xml = remover_secoes_ausentes(xml, decisao)
                xml = limpar_marcadores_secao(xml)

                num_atividades = len(inspecoes)
                if num_atividades >= 1:
                    xml = ajustar_linhas_quadro(xml, num_atividades)

                num_ncs = len(dados.get("nao_conformidades", []))
                if num_ncs == 0:
                    num_ncs = 3
                xml = ajustar_linhas_nc(xml, num_ncs)

                mapa = construir_mapa(dados)
                xml = aplicar_substituicoes(xml, mapa)
                # OBS: coloração condicional das células de Prioridade (NCs)
                # e Conformidade (Eficiência). Roda DEPOIS das substituições
                # pra ter os valores finais.
                xml = aplicar_cores_celulas(xml, mapa)
                doc_xml_path.write_text(xml, encoding="utf-8")

                for nome_arq in ["header1.xml", "header2.xml", "footer1.xml", "footer2.xml"]:
                    arq_path = unpacked / "word" / nome_arq
                    if arq_path.exists():
                        xml_aux = arq_path.read_text(encoding="utf-8")
                        xml_aux = aplicar_substituicoes(xml_aux, mapa)
                        arq_path.write_text(xml_aux, encoding="utf-8")

                if imagens_a_processar:
                    copiar_fotos_para_media(unpacked, imagens_a_processar)
                    adicionar_relationships_imagens(unpacked, imagens_a_processar)

            # Recompactar como .docx (comum aos dois fluxos)
            output_path = tmp / nome_arquivo
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z_out:
                for arq in unpacked.rglob("*"):
                    if arq.is_file():
                        z_out.write(arq, arq.relative_to(unpacked))

            return send_file(
                output_path,
                as_attachment=True,
                download_name=nome_arquivo,
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500


if __name__ == "__main__":
    # Pra rodar local. No Render, use `gunicorn app:app`
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
