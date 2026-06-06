"""
secao4_builder.py — Constrói o XML dinâmico da Seção 4 (Detalhamento das Atividades)
e do trecho 4.x do sumário, com base nos dados das inspeções do mês.

A geração programática preserva o estilo visual do espelho:
- Título da seção em vermelho (Título 1)
- Subtítulos 4.x em vermelho (Título 2)
- Texto justificado, fonte e espaçamento padrão
- Imagens centralizadas com legendas ABNT abaixo (itálico cinza)
- Múltiplas fotos: lado a lado em tabela invisível (~8x6cm cada)
"""

import shutil
import unicodedata
from pathlib import Path
from typing import List, Dict, Any

# Pillow para ler dimensões reais das imagens
try:
    from PIL import Image
    PILLOW_OK = True
except ImportError:
    PILLOW_OK = False


# ============================================================
# CONSTANTES
# ============================================================

EMU_POR_POLEGADA = 914400

# Fotos
FOTO_UNICA_LARGURA_POL = 3.8
FOTO_UNICA_ALTURA_POL = 4.0

# Para fotos lado a lado: alvo ~8x6cm = 3.15" x 2.36"
FOTO_LADO_LADO_LARGURA_POL = 3.15
FOTO_LADO_LADO_ALTURA_POL = 2.36

# Largura útil da página (A4 menos margens)
PAGINA_LARGURA_DXA = 9000  # twentieths of a point


def escape_xml(t: str) -> str:
    if t is None:
        return ""
    return (str(t)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;"))


def calcular_emu(largura_px: int, altura_px: int, max_l_pol: float, max_a_pol: float) -> tuple:
    if largura_px <= 0 or altura_px <= 0:
        return int(max_l_pol * EMU_POR_POLEGADA), int(max_a_pol * EMU_POR_POLEGADA)
    proporcao = largura_px / altura_px
    largura_pol = max_l_pol
    altura_pol = largura_pol / proporcao
    if altura_pol > max_a_pol:
        altura_pol = max_a_pol
        largura_pol = altura_pol * proporcao
    return int(largura_pol * EMU_POR_POLEGADA), int(altura_pol * EMU_POR_POLEGADA)


def ler_dimensoes_px(caminho: Path) -> tuple:
    if not PILLOW_OK:
        return (0, 0)
    try:
        with Image.open(caminho) as img:
            return img.size
    except Exception:
        return (0, 0)


# ============================================================
# CONSTRUÇÃO DE XML — BLOCOS BÁSICOS
# ============================================================

def p_titulo_secao4() -> str:
    """Parágrafo do título '4. DETALHAMENTO DAS ATIVIDADES' (Verdana 16pt navy negrito).

    NÃO é chamado por construir_secao4() — o título já vem do template.
    Mantido apenas por compatibilidade.
    """
    return '''<w:p>
        <w:pPr>
          <w:spacing w:before="360" w:after="180" w:line="360" w:lineRule="auto"/>
          <w:jc w:val="left"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Verdana" w:cs="Verdana" w:eastAsia="Verdana" w:hAnsi="Verdana"/>
            <w:b/>
            <w:bCs/>
            <w:color w:val="0A2540"/>
            <w:sz w:val="32"/>
            <w:szCs w:val="32"/>
          </w:rPr>
          <w:t>4. DETALHAMENTO DAS ATIVIDADES</w:t>
        </w:r>
      </w:p>'''


def p_subtitulo(num: str, titulo: str) -> str:
    """Subtítulo '4.X. Título' em Verdana negrito 12pt navy."""
    return f'''<w:p>
        <w:pPr>
          <w:spacing w:before="280" w:after="120" w:line="360" w:lineRule="auto"/>
          <w:jc w:val="left"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Verdana" w:cs="Verdana" w:eastAsia="Verdana" w:hAnsi="Verdana"/>
            <w:b/>
            <w:bCs/>
            <w:color w:val="0A2540"/>
            <w:sz w:val="24"/>
            <w:szCs w:val="24"/>
          </w:rPr>
          <w:t>{escape_xml(num)}. {escape_xml(titulo)}</w:t>
        </w:r>
      </w:p>'''


def p_texto(texto: str, negrito_inicio: str = None) -> str:
    """
    Parágrafo de texto justificado em padrão ABNT (Verdana 12pt, 1,5 entrelinhas,
    recuo de 1,25 cm na primeira linha). Se negrito_inicio for fornecido
    (ex.: 'Medida corretiva recomendada:'), o início aparece em negrito.
    """
    fonte_rpr = (
        '<w:rFonts w:ascii="Verdana" w:cs="Verdana" w:eastAsia="Verdana" w:hAnsi="Verdana"/>'
        '<w:sz w:val="22"/><w:szCs w:val="22"/>'
    )
    runs = []
    if negrito_inicio:
        runs.append(
            f'<w:r><w:rPr>{fonte_rpr}<w:b/><w:bCs/></w:rPr>'
            f'<w:t xml:space="preserve">{escape_xml(negrito_inicio)} </w:t></w:r>'
        )
    runs.append(
        f'<w:r><w:rPr>{fonte_rpr}</w:rPr>'
        f'<w:t xml:space="preserve">{escape_xml(texto)}</w:t></w:r>'
    )
    return f'''<w:p>
        <w:pPr>
          <w:spacing w:before="0" w:after="120" w:line="360" w:lineRule="auto"/>
          <w:ind w:firstLine="708"/>
          <w:jc w:val="both"/>
        </w:pPr>
        {"".join(runs)}
      </w:p>'''


def p_legenda(texto: str) -> str:
    """Legenda da figura, centralizada e em itálico cinza (Verdana 10pt)."""
    return f'''<w:p>
        <w:pPr>
          <w:spacing w:before="0" w:after="240" w:line="276" w:lineRule="auto"/>
          <w:jc w:val="center"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Verdana" w:cs="Verdana" w:eastAsia="Verdana" w:hAnsi="Verdana"/>
            <w:i/><w:iCs/>
            <w:color w:val="666666"/>
            <w:sz w:val="20"/><w:szCs w:val="20"/>
          </w:rPr>
          <w:t>{escape_xml(texto)}</w:t>
        </w:r>
      </w:p>'''


def drawing_imagem(rid: str, cx: int, cy: int, nome: str = "Foto", id_img: int = 1) -> str:
    """Bloco <w:drawing> para uma imagem inline com tamanho dado em EMU."""
    return f'''<w:r><w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="{cx}" cy="{cy}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="{id_img}" name="{escape_xml(nome)}"/>
        <wp:cNvGraphicFramePr>
          <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
        </wp:cNvGraphicFramePr>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:nvPicPr>
                <pic:cNvPr id="{id_img}" name="{escape_xml(nome)}"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{rid}"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing></w:r>'''


def p_imagem_centralizada(rid: str, cx: int, cy: int, nome: str, id_img: int) -> str:
    """Parágrafo com uma imagem centralizada."""
    return f'''<w:p>
        <w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="60"/></w:pPr>
        {drawing_imagem(rid, cx, cy, nome, id_img)}
      </w:p>'''


def tabela_imagens_lado_a_lado(imagens: list) -> str:
    """
    imagens = [{"rid": str, "cx": int, "cy": int, "nome": str, "id": int, "legenda": str}, ...]
    Quebra em linhas de 2 imagens cada (ou 1 se for só uma).
    Tabela invisível para alinhamento.
    """
    if not imagens:
        return ""

    # Quebrar em pares
    linhas_pares = [imagens[i:i+2] for i in range(0, len(imagens), 2)]

    largura_celula = PAGINA_LARGURA_DXA // 2

    linhas_xml = []
    for par in linhas_pares:
        celulas = []
        for img in par:
            celulas.append(f'''<w:tc>
                <w:tcPr>
                  <w:tcW w:w="{largura_celula}" w:type="dxa"/>
                  <w:tcBorders>
                    <w:top w:val="nil"/><w:bottom w:val="nil"/>
                    <w:left w:val="nil"/><w:right w:val="nil"/>
                  </w:tcBorders>
                </w:tcPr>
                <w:p>
                  <w:pPr><w:jc w:val="center"/><w:spacing w:before="60" w:after="0"/></w:pPr>
                  {drawing_imagem(img["rid"], img["cx"], img["cy"], img["nome"], img["id"])}
                </w:p>
                <w:p>
                  <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="240"/></w:pPr>
                  <w:r>
                    <w:rPr>
                      <w:rFonts w:ascii="Verdana" w:cs="Verdana" w:eastAsia="Verdana" w:hAnsi="Verdana"/>
                      <w:i/><w:iCs/>
                      <w:color w:val="666666"/>
                      <w:sz w:val="20"/><w:szCs w:val="20"/>
                    </w:rPr>
                    <w:t>{escape_xml(img["legenda"])}</w:t>
                  </w:r>
                </w:p>
              </w:tc>''')

        # Se for par incompleto (1 só), adiciona célula vazia
        if len(par) == 1:
            celulas.append(f'''<w:tc>
                <w:tcPr>
                  <w:tcW w:w="{largura_celula}" w:type="dxa"/>
                  <w:tcBorders>
                    <w:top w:val="nil"/><w:bottom w:val="nil"/>
                    <w:left w:val="nil"/><w:right w:val="nil"/>
                  </w:tcBorders>
                </w:tcPr>
                <w:p/>
              </w:tc>''')

        linhas_xml.append(f'<w:tr>{"".join(celulas)}</w:tr>')

    return f'''<w:tbl>
      <w:tblPr>
        <w:tblW w:w="{PAGINA_LARGURA_DXA}" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="nil"/><w:bottom w:val="nil"/>
          <w:left w:val="nil"/><w:right w:val="nil"/>
          <w:insideH w:val="nil"/><w:insideV w:val="nil"/>
        </w:tblBorders>
        <w:jc w:val="center"/>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="{largura_celula}"/>
        <w:gridCol w:w="{largura_celula}"/>
      </w:tblGrid>
      {"".join(linhas_xml)}
    </w:tbl>'''


# ============================================================
# CONSTRUÇÃO DA SEÇÃO 4 INTEIRA
# ============================================================

def construir_secao4(
    inspecoes: List[Dict[str, Any]],
    pasta_assets: Path,
    pasta_unpacked: Path,
    rids_disponiveis: List[str],
) -> tuple:
    """
    Gera o XML completo da Seção 4 a partir da lista de inspeções e
    retorna (xml_secao4, lista_imagens_a_copiar).

    Cada inspeção:
        {
          "titulo": "Inspeção no Painel Elétrico",
          "achados": "texto livre" ou None,
          "medida": "texto livre" ou None,
          "fotos": [
            {"arquivo": "path/foto1.jpg", "legenda": "..."},
            ...
          ]
        }

    rids_disponiveis: lista de rIds reservados para imagens novas
    (imagem9, imagem10, ... — separados dos rIds do template)

    Retorna também a lista [(arquivo_origem, nome_destino, rid)] para o caller
    copiar os arquivos e gerar as relações.
    """
    blocos = []  # Título já vem do template
    imagens_a_processar = []  # (origem, nome_destino, rid)
    rid_idx = 0
    figura_global = 1
    img_id_global = 100  # ids únicos para wp:docPr

    for i, insp in enumerate(inspecoes, start=1):
        # Subtítulo 4.X
        blocos.append(p_subtitulo(f"4.{i}", insp.get("titulo", f"Inspeção {i}")))

        # Descrição em parágrafos (formato novo: lista de strings)
        descricao_paragrafos = insp.get("descricao_paragrafos")
        if descricao_paragrafos:
            for paragrafo in descricao_paragrafos:
                texto = paragrafo.strip() if isinstance(paragrafo, str) else ""
                if not texto:
                    continue
                # Se o parágrafo começa com "Medida corretiva..." aplica negrito
                if texto.lower().startswith("medida corretiva"):
                    # Detectar o ponto onde acaba a parte em negrito
                    idx = texto.find(":")
                    if idx > 0:
                        blocos.append(p_texto(texto[idx+1:].strip(),
                                              negrito_inicio=texto[:idx+1]))
                    else:
                        blocos.append(p_texto(texto))
                else:
                    blocos.append(p_texto(texto))
        else:
            # Formato antigo (compatibilidade)
            achados = insp.get("achados")
            if achados:
                for paragrafo in achados.split("\n\n"):
                    if paragrafo.strip():
                        blocos.append(p_texto(paragrafo.strip()))
            medida = insp.get("medida")
            if medida:
                blocos.append(p_texto(medida.strip(),
                                      negrito_inicio="Medida corretiva recomendada:"))

        # Fotos
        fotos = insp.get("fotos", [])
        if fotos:
            imagens_dados = []
            for foto in fotos:
                if rid_idx >= len(rids_disponiveis):
                    print(f"⚠ Sem rIds suficientes — limite de {len(rids_disponiveis)} fotos.")
                    break

                rid = rids_disponiveis[rid_idx]
                rid_idx += 1

                arquivo_orig = Path(foto["arquivo"])
                if not arquivo_orig.is_absolute():
                    arquivo_orig = pasta_assets / arquivo_orig

                if not arquivo_orig.exists():
                    print(f"⚠ Foto não encontrada: {arquivo_orig}")
                    continue

                # Nome destino preserva extensão
                ext = arquivo_orig.suffix.lower()
                nome_destino = f"image_dyn_{rid_idx}{ext}"

                largura_px, altura_px = ler_dimensoes_px(arquivo_orig)

                # Decide tamanho: se a inspeção tem 1 foto, tamanho cheio.
                # Se tem 2+, usa o tamanho lado-a-lado (~8x6cm cada).
                if len(fotos) == 1:
                    cx, cy = calcular_emu(largura_px, altura_px,
                                          FOTO_UNICA_LARGURA_POL, FOTO_UNICA_ALTURA_POL)
                else:
                    cx, cy = calcular_emu(largura_px, altura_px,
                                          FOTO_LADO_LADO_LARGURA_POL, FOTO_LADO_LADO_ALTURA_POL)

                imagens_a_processar.append((arquivo_orig, nome_destino, rid))

                legenda_texto = foto.get("legenda", "").strip()
                legenda_completa = f"Figura {figura_global} – {legenda_texto}" if legenda_texto else f"Figura {figura_global}"

                imagens_dados.append({
                    "rid": rid,
                    "cx": cx,
                    "cy": cy,
                    "nome": f"Figura {figura_global}",
                    "id": img_id_global,
                    "legenda": legenda_completa,
                })
                figura_global += 1
                img_id_global += 1

            # Renderizar fotos
            if len(imagens_dados) == 1:
                img = imagens_dados[0]
                blocos.append(p_imagem_centralizada(img["rid"], img["cx"], img["cy"], img["nome"], img["id"]))
                blocos.append(p_legenda(img["legenda"]))
            elif len(imagens_dados) > 1:
                blocos.append(tabela_imagens_lado_a_lado(imagens_dados))

    return "".join(blocos), imagens_a_processar


def construir_sumario_4x(inspecoes: List[Dict[str, Any]]) -> str:
    """Gera os parágrafos do sumário para as subseções 4.x.

    Estilo: recuo à esquerda, número 4.X em coluna, título do subitem em
    Title Case PT-BR (sem negrito), pontilhados até a margem direita.
    Sem número de página por enquanto.
    """
    # Palavras curtas que devem ficar minúsculas em pt-BR (exceto se forem
    # a primeira palavra do título)
    MINUSCULAS = {
        "a", "as", "à", "às", "ao", "aos", "o", "os", "um", "uma", "uns", "umas",
        "de", "do", "da", "dos", "das", "em", "no", "na", "nos", "nas",
        "por", "para", "pela", "pelo", "pelas", "pelos",
        "com", "sem", "sob", "sobre", "entre", "ante", "após",
        "e", "ou", "mas", "que", "se",
    }

    def title_case_ptbr(texto: str) -> str:
        palavras = texto.strip().split()
        resultado = []
        for i, p in enumerate(palavras):
            p_lower = p.lower()
            if i == 0 or p_lower not in MINUSCULAS:
                # Primeira palavra ou palavra "longa": capitalizar
                resultado.append(p_lower.capitalize())
            else:
                resultado.append(p_lower)
        return " ".join(resultado)

    blocos = []
    for i, insp in enumerate(inspecoes, start=1):
        titulo_raw = insp.get("titulo", f"Inspeção {i}")
        titulo_tc = title_case_ptbr(titulo_raw) if titulo_raw else f"Inspeção {i}"
        titulo = escape_xml(titulo_tc)
        blocos.append(
            '<w:p>'
            '<w:pPr>'
            '<w:tabs>'
            '<w:tab w:val="left" w:pos="2160"/>'
            '<w:tab w:val="right" w:pos="9360" w:leader="dot"/>'
            '</w:tabs>'
            '<w:spacing w:after="120" w:before="120"/>'
            '<w:ind w:left="1440"/>'
            '</w:pPr>'
            '<w:r>'
            '<w:rPr>'
            '<w:rFonts w:ascii="Verdana" w:cs="Verdana" w:eastAsia="Verdana" w:hAnsi="Verdana"/>'
            '<w:color w:val="000000"/>'
            '<w:sz w:val="24"/>'
            '<w:szCs w:val="24"/>'
            '</w:rPr>'
            f'<w:t xml:space="preserve">4.{i}.\t{titulo}</w:t>'
            '</w:r>'
            '</w:p>'
        )
    return "".join(blocos)


# ============================================================
# ANEXO DO LAUDO (PDF → imagens inseridas no relatório)
# ============================================================

# Tamanho de cada página do laudo no documento (A4 útil, retrato)
LAUDO_LARGURA_POL = 6.0   # largura da imagem da página do laudo
LAUDO_ALTURA_POL = 9.2    # altura máxima (cabe numa página A4 com margens)


def construir_anexo_laudo(
    laudo_url: str,
    pasta_assets: Path,
    rids_disponiveis: List[str],
) -> tuple:
    """
    Baixa o PDF do laudo a partir de laudo_url, converte cada página em imagem
    (JPEG) e gera o XML para inserir todas as páginas no relatório, uma por
    parágrafo centralizado.

    Retorna (xml_paginas, lista_imagens_a_copiar), onde lista_imagens é
    [(arquivo_origem, nome_destino, rid), ...] no mesmo formato que a Seção 4
    usa — para o caller copiar pra media/ e criar os Relationships.

    Se algo falhar (download, PyMuPDF ausente, PDF inválido), retorna ("", [])
    para o relatório ser gerado sem o anexo (degradação graciosa).
    """
    if not laudo_url:
        return "", []

    try:
        import fitz  # PyMuPDF
    except ImportError:
        print("⚠ PyMuPDF (fitz) não instalado — anexo do laudo ignorado")
        return "", []

    # 1) Baixar o PDF
    pdf_path = pasta_assets / "laudo_baixado.pdf"
    try:
        import urllib.request
        urllib.request.urlretrieve(laudo_url, pdf_path)
    except Exception as e:
        print(f"⚠ Falha ao baixar laudo de {laudo_url}: {e}")
        return "", []

    # 2) Converter páginas em imagens JPEG
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"⚠ PDF do laudo inválido: {e}")
        return "", []

    blocos = []
    imagens = []
    mat = fitz.Matrix(1.5, 1.5)  # ~108 DPI, bom equilíbrio leitura/tamanho
    img_id = 500  # ids únicos pra wp:docPr (longe dos da seção 4)

    n_paginas = len(doc)
    for i in range(n_paginas):
        if i >= len(rids_disponiveis):
            print(f"⚠ rIds insuficientes para o laudo (parou na página {i+1})")
            break
        rid = rids_disponiveis[i]
        nome_destino = f"laudo_p{i+1}.jpg"
        arquivo = pasta_assets / nome_destino
        try:
            pix = doc[i].get_pixmap(matrix=mat)
            pix.save(arquivo, jpg_quality=80)
        except Exception as e:
            print(f"⚠ Falha ao converter página {i+1} do laudo: {e}")
            continue

        # Dimensões reais pra manter proporção
        larg_px, alt_px = ler_dimensoes_px(arquivo)
        cx, cy = calcular_emu(larg_px, alt_px, LAUDO_LARGURA_POL, LAUDO_ALTURA_POL)

        blocos.append(p_imagem_centralizada(rid, cx, cy, f"Laudo página {i+1}", img_id))
        img_id += 1
        imagens.append((str(arquivo), nome_destino, rid))

    doc.close()
    return "".join(blocos), imagens
