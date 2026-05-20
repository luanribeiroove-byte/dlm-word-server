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
    decidir_secoes_e_numerar, remover_secoes_ausentes, limpar_marcadores_secao
)
from secao4_builder import construir_secao4, construir_sumario_4x
import zipfile

app = Flask(__name__)
CORS(app)  # Permite chamadas do app web em qualquer origem

ROOT = Path(__file__).parent
TEMPLATE = ROOT / "template.docx"


@app.route("/", methods=["GET"])
def health():
    """Health check. O Render usa esse endpoint pra ver se o serviço tá vivo."""
    return jsonify({
        "ok": True,
        "service": "dlm-word-generator",
        "version": "1.0.0",
        "template_exists": TEMPLATE.exists()
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

        # Nome do arquivo de saída
        cliente = dados.get("cabecalho", {}).get("cliente", "Relatorio")
        periodo = dados.get("cabecalho", {}).get("periodo", "")
        nome_arquivo = f"Relatorio_{cliente.replace(' ', '_')}_{periodo.replace('/', '_')}.docx"
        # Sanitizar
        nome_arquivo = "".join(c if c.isalnum() or c in "._-" else "_" for c in nome_arquivo)

        # Gerar em pasta temporária
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            unpacked = tmp / "unpacked"

            # Extrair template
            with zipfile.ZipFile(TEMPLATE, "r") as z:
                z.extractall(unpacked)

            doc_xml_path = unpacked / "word" / "document.xml"
            xml = doc_xml_path.read_text(encoding="utf-8")

            # Seção 4 dinâmica
            inspecoes = dados.get("inspecoes", [])
            if not inspecoes:
                secao4_xml = ""
                sumario_4x_xml = ""
                imagens_a_processar = []
            else:
                total_fotos = sum(len(i.get("fotos", [])) for i in inspecoes)
                rids_disponiveis = reservar_rids_para_fotos(unpacked, total_fotos) if total_fotos else []
                # pasta_assets = tmp (sem fotos no MVP)
                secao4_xml, imagens_a_processar = construir_secao4(
                    inspecoes, tmp, unpacked, rids_disponiveis
                )
                sumario_4x_xml = construir_sumario_4x(inspecoes)

            # Substituir áreas marcadas
            # Substituir áreas marcadas (Seção 4 dinâmica)
            xml = substituir_area_marcada(xml, "__SUMARIO_4X_INICIO__", "__SUMARIO_4X_FIM__", sumario_4x_xml)
            xml = substituir_area_marcada(xml, "__SECAO4_INICIO__", "__SECAO4_FIM__", secao4_xml)

            # ===== Remover seções ausentes e limpar marcadores das presentes =====
            decisao = decidir_secoes_e_numerar(dados)
            xml = remover_secoes_ausentes(xml, decisao)
            xml = limpar_marcadores_secao(xml)

            # Ajustes dinâmicos
            num_atividades = len(inspecoes)
            if num_atividades >= 1:
                xml = ajustar_linhas_quadro(xml, num_atividades)

            num_ncs = len(dados.get("nao_conformidades", []))
            if num_ncs == 0:
                num_ncs = 3
            xml = ajustar_linhas_nc(xml, num_ncs)

            # Substituir placeholders
            mapa = construir_mapa(dados)
            xml = aplicar_substituicoes(xml, mapa)

            # Salvar XML modificado
            doc_xml_path.write_text(xml, encoding="utf-8")

            # Aplicar substituições também no header e footer (que contêm placeholders do cliente)
            for nome_arq in ["header1.xml", "header2.xml", "footer1.xml", "footer2.xml"]:
                arq_path = unpacked / "word" / nome_arq
                if arq_path.exists():
                    xml_aux = arq_path.read_text(encoding="utf-8")
                    xml_aux = aplicar_substituicoes(xml_aux, mapa)
                    arq_path.write_text(xml_aux, encoding="utf-8")

            # Copiar fotos (se houver — no MVP geralmente não)
            if imagens_a_processar:
                copiar_fotos_para_media(unpacked, imagens_a_processar)
                adicionar_relationships_imagens(unpacked, imagens_a_processar)

            # Recompactar como .docx
            output_path = tmp / nome_arquivo
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z_out:
                for arq in unpacked.rglob("*"):
                    if arq.is_file():
                        z_out.write(arq, arq.relative_to(unpacked))

            # Enviar
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
