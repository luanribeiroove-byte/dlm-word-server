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
    """Parágrafo do título '4. DETALHAMENTO DAS ATIVIDADES' em vermelho."""
    return '''<w:p>
        <w:pPr>
          <w:pStyle w:val="Ttulo1"/>
          <w:rPr><w:color w:val="000000"/></w:rPr>
        </w:pPr>
        <w:r>
          <w:rPr><w:color w:val="000000"/></w:rPr>
          <w:t>4. DETALHAMENTO DAS ATIVIDADES</w:t>
        </w:r>
      </w:p>'''


def p_subtitulo(num: str, titulo: str) -> str:
    """Subtítulo '4.X. Título' em navy com negrito e tamanho maior."""
    return f'''<w:p>
        <w:pPr>
          <w:spacing w:before="240" w:after="120"/>
          <w:keepNext/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Arial" w:cs="Arial" w:eastAsia="Arial" w:hAnsi="Arial"/>
            <w:b/><w:bCs/>
            <w:color w:val="0A2540"/>
            <w:sz w:val="26"/><w:szCs w:val="26"/>
          </w:rPr>
          <w:t>{escape_xml(num)}. {escape_xml(titulo)}</w:t>
        </w:r>
      </w:p>'''


def p_texto(texto: str, negrito_inicio: str = None) -> str:
    """
    Parágrafo de texto justificado. Se negrito_inicio for fornecido (ex.:
    'Medida corretiva recomendada:'), gera DOIS parágrafos: um com o título em
    negrito e outro com o texto abaixo.
    """
    if negrito_inicio:
        # Parágrafo do título (em negrito, sem espaço grande embaixo)
        titulo_xml = f'''<w:p>
            <w:pPr>
              <w:spacing w:before="120" w:after="0" w:line="360" w:lineRule="auto"/>
              <w:keepNext/>
            </w:pPr>
            <w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">{escape_xml(negrito_inicio)}</w:t></w:r>
          </w:p>'''
        # Parágrafo do texto (logo abaixo, sem espaço extra em cima)
        texto_xml = f'''<w:p>
            <w:pPr>
              <w:spacing w:before="0" w:after="80" w:line="360" w:lineRule="auto"/>
              <w:jc w:val="both"/>
            </w:pPr>
            <w:r><w:t xml:space="preserve">{escape_xml(texto)}</w:t></w:r>
          </w:p>'''
        return titulo_xml + texto_xml
    # Caso normal: parágrafo único
    return f'''<w:p>
        <w:pPr>
          <w:spacing w:before="80" w:after="80" w:line="360" w:lineRule="auto"/>
          <w:jc w:val="both"/>
        </w:pPr>
        <w:r><w:t xml:space="preserve">{escape_xml(texto)}</w:t></w:r>
      </w:p>'''


def p_legenda(texto: str) -> str:
    """Legenda da figura, centralizada e em itálico cinza."""
    return f'''<w:p>
        <w:pPr>
          <w:spacing w:before="0" w:after="240"/>
          <w:jc w:val="center"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:i/><w:iCs/>
            <w:color w:val="666666"/>
            <w:sz w:val="18"/><w:szCs w:val="18"/>
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
                    <w:rPr><w:i/><w:iCs/><w:color w:val="666666"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
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
    """Gera os parágrafos do sumário para as subseções 4.x."""
    blocos = []
    for i, insp in enumerate(inspecoes, start=1):
        titulo = escape_xml(insp.get("titulo", f"Inspeção {i}"))
        blocos.append(f'''<w:p>
        <w:pPr>
          <w:pStyle w:val="Sumrio2"/>
          <w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9016"/></w:tabs>
          <w:rPr>
            <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
            <w:noProof/>
          </w:rPr>
        </w:pPr>
        <w:r><w:rPr><w:noProof/></w:rPr><w:t>4.{i}. {titulo}</w:t></w:r>
      </w:p>''')
    return "".join(blocos)
