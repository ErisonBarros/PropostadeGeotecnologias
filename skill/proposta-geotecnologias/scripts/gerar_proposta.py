#!/usr/bin/env python3
"""Gera a proposta em DOCX (folha modelo R. Oliveira), PDF e Markdown.

Uso:
    python3 gerar_proposta.py dados.yaml [--saida DIR] [--sem-pdf]

Saídas (em --saida, padrão ./saida_proposta):
    AAAA_MM_DD_vNN.md                      cópia de arquivo (padrão exigido)
    Proposta_<num>_<slug>_AAAA-MM-DD_vNN.docx
    Proposta_<num>_<slug>_AAAA-MM-DD_vNN.pdf   (se LibreOffice disponível)
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import _comum as C  # noqa: E402

C.garantir_dependencias()

from docx import Document  # noqa: E402
from docx.enum.section import WD_SECTION  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt, RGBColor  # noqa: E402
from lxml import etree  # noqa: E402

# ------------------------------------------------------------------ identidade
AZUL = "050A30"        # faixa do cabeçalho / títulos
AZUL2 = "03406D"       # faixa inferior / realces
GRAFITE = "2B2B2B"
CINZA = "5A6472"
BORDA = "C3CBD6"
FUNDO_CLARO = "F2F4F7"
FUNDO_AZUL = "DCE6F1"
FONTE = "Calibri"


# ------------------------------------------------------------------ helpers XML
def _cel_fundo(cel, cor: str) -> None:
    tcPr = cel._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), cor)
    tcPr.append(shd)


def _cel_margens(cel, top=70, bottom=70, left=110, right=110) -> None:
    tcPr = cel._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for lado, valor in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{lado}")
        node.set(qn("w:w"), str(valor))
        node.set(qn("w:type"), "dxa")
        mar.append(node)
    tcPr.append(mar)


def _tabela_bordas(tabela, cor: str = BORDA, sz: int = 4) -> None:
    tblPr = tabela._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), cor)
        borders.append(el)
    tblPr.append(borders)


def _sem_bordas(tabela) -> None:
    tblPr = tabela._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)


def _repetir_cabecalho(linha) -> None:
    trPr = linha._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def _nao_quebrar(linha) -> None:
    trPr = linha._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def _borda_paragrafo(p, lado="bottom", cor=AZUL2, sz=6, espaco=4) -> None:
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    el = OxmlElement(f"w:{lado}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(espaco))
    el.set(qn("w:color"), cor)
    pBdr.append(el)
    pPr.append(pBdr)


def _campo(par, instrucao: str) -> None:
    run = par.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instrucao
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(fld_end)


def _imagem_pagina(par, caminho: Path, largura_cm: float, y_cm: float, atras=True) -> None:
    """Insere imagem ancorada na página (sangria total), independente das margens."""
    run = par.add_run()
    run.add_picture(str(caminho), width=Cm(largura_cm))
    drawing = run._r.find(qn("w:drawing"))
    inline = drawing[0]

    anchor = OxmlElement("wp:anchor")
    for chave, valor in (("distT", "0"), ("distB", "0"), ("distL", "0"), ("distR", "0"),
                         ("simplePos", "0"), ("relativeHeight", "2"),
                         ("behindDoc", "1" if atras else "0"), ("locked", "0"),
                         ("layoutInCell", "1"), ("allowOverlap", "1")):
        anchor.set(chave, valor)

    simple = OxmlElement("wp:simplePos")
    simple.set("x", "0")
    simple.set("y", "0")
    anchor.append(simple)

    for eixo, valor in (("H", 0), ("V", int(Cm(y_cm)))):
        pos = OxmlElement(f"wp:position{eixo}")
        pos.set("relativeFrom", "page")
        off = OxmlElement("wp:posOffset")
        off.text = str(valor)
        pos.append(off)
        anchor.append(pos)

    for filho in list(inline):
        nome = etree.QName(filho).localname
        anchor.append(filho)
        if nome == "extent":
            eff = OxmlElement("wp:effectExtent")
            for k in ("l", "t", "r", "b"):
                eff.set(k, "0")
            anchor.append(eff)
            anchor.append(OxmlElement("wp:wrapNone"))

    drawing.remove(inline)
    drawing.append(anchor)


def _larguras(tabela, larguras_cm) -> None:
    """Fixa a largura das colunas (tblW + gridCol), respeitada por Word e LibreOffice."""
    tbl = tabela._tbl
    tblPr = tbl.tblPr
    total = int(round(sum(larguras_cm) * 566.93))
    for el in tblPr.findall(qn("w:tblW")):
        tblPr.remove(el)
    larg = OxmlElement("w:tblW")
    larg.set(qn("w:type"), "dxa")
    larg.set(qn("w:w"), str(total))
    tblPr.insert(0, larg)
    for el in tblPr.findall(qn("w:tblLayout")):
        tblPr.remove(el)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)

    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for gc in list(grid):
            grid.remove(gc)
        for valor in larguras_cm:
            gc = OxmlElement("w:gridCol")
            gc.set(qn("w:w"), str(int(round(valor * 566.93))))
            grid.append(gc)
    for row in tabela.rows:
        for i, cel in enumerate(row.cells):
            if i < len(larguras_cm):
                cel.width = Cm(larguras_cm[i])


# ------------------------------------------------------------------ documento
def montar_folha(doc) -> None:
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin = Cm(4.0)
    sec.bottom_margin = Cm(3.35)
    sec.header_distance = Cm(0)
    sec.footer_distance = Cm(2.35)

    estilo = doc.styles["Normal"]
    estilo.font.name = FONTE
    estilo.font.size = Pt(10.5)
    estilo.font.color.rgb = RGBColor.from_string(GRAFITE)
    estilo.paragraph_format.space_after = Pt(6)
    estilo.paragraph_format.line_spacing = 1.12
    rpr = estilo.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONTE)

    # ---- cabeçalho: faixa com logotipo, ancorada na página (sangria total)
    hdr = sec.header
    hdr.is_linked_to_previous = False
    p = hdr.paragraphs[0]
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1
    p.add_run().font.size = Pt(1)
    _imagem_pagina(p, C.IDENTIDADE / "cabecalho.png", 21.0, 0.0)

    # ---- rodapé: numeração acima da faixa de contato
    ftr = sec.footer
    ftr.is_linked_to_previous = False
    pnum = ftr.paragraphs[0]
    pnum.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pnum.paragraph_format.space_after = Pt(0)
    pnum.paragraph_format.space_before = Pt(0)
    r = pnum.add_run("Página ")
    r.font.size = Pt(7.5)
    r.font.color.rgb = RGBColor.from_string(CINZA)
    _campo(pnum, "PAGE")
    r = pnum.add_run(" de ")
    r.font.size = Pt(7.5)
    r.font.color.rgb = RGBColor.from_string(CINZA)
    _campo(pnum, "NUMPAGES")
    for run in pnum.runs:
        run.font.size = Pt(7.5)
        run.font.color.rgb = RGBColor.from_string(CINZA)

    altura_faixa = 21.0 * 135 / 1300  # proporção da arte de rodapé
    pimg = ftr.add_paragraph()
    pf = pimg.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1
    pimg.add_run().font.size = Pt(1)
    _imagem_pagina(pimg, C.IDENTIDADE / "rodape.png", 21.0, 29.7 - altura_faixa)


def titulo_secao(doc, numero: str, texto: str) -> None:
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(14)
    pf.space_after = Pt(8)
    pf.keep_with_next = True
    run = p.add_run(f"{numero}. {texto.upper()}")
    run.bold = True
    run.font.size = Pt(11.5)
    run.font.color.rgb = RGBColor.from_string(AZUL)
    run.font.name = FONTE
    _borda_paragrafo(p, "bottom", AZUL2, 6, 3)


def sub(doc, texto: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(texto)
    run.bold = True
    run.font.size = Pt(10.5)
    run.font.color.rgb = RGBColor.from_string(AZUL2)


def par(doc, texto: str = "", negrito=False, italico=False, tamanho=10.5,
        alinhamento="just", espaco=6, cor=GRAFITE) -> None:
    p = doc.add_paragraph()
    p.alignment = {"just": WD_ALIGN_PARAGRAPH.JUSTIFY, "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "left": WD_ALIGN_PARAGRAPH.LEFT, "right": WD_ALIGN_PARAGRAPH.RIGHT}[alinhamento]
    p.paragraph_format.space_after = Pt(espaco)
    run = p.add_run(texto)
    run.bold = negrito
    run.italic = italico
    run.font.size = Pt(tamanho)
    run.font.color.rgb = RGBColor.from_string(cor)


def bullets(doc, itens, marcador="•", tamanho=10.5) -> None:
    for item in itens:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Cm(0.75)
        pf.first_line_indent = Cm(-0.35)
        pf.space_after = Pt(3)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = p.add_run(f"{marcador}  ")
        run.font.size = Pt(tamanho)
        run.font.color.rgb = RGBColor.from_string(AZUL2)
        run.bold = True
        # permite negrito inline com **texto**
        texto = str(item)
        partes = texto.split("**")
        for i, parte in enumerate(partes):
            r = p.add_run(parte)
            r.font.size = Pt(tamanho)
            r.bold = i % 2 == 1


def tabela(doc, cabecalhos, linhas, larguras, alinhamentos=None, linha_total=False,
           tam=9.5, cabecalho_visivel=True) -> None:
    ncols = len(larguras)
    t = doc.add_table(rows=0, cols=ncols)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    _tabela_bordas(t)
    alinhamentos = alinhamentos or ["left"] * ncols
    mapa = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT, "just": WD_ALIGN_PARAGRAPH.JUSTIFY}

    def escrever(cel, texto, negrito=False, cor=GRAFITE, alinh="left", tamanho=tam):
        cel.text = ""
        p = cel.paragraphs[0]
        p.alignment = mapa[alinh]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(str(texto))
        r.bold = negrito
        r.font.size = Pt(tamanho)
        r.font.color.rgb = RGBColor.from_string(cor)
        r.font.name = FONTE

    if cabecalho_visivel and cabecalhos:
        row = t.add_row()
        _repetir_cabecalho(row)
        _nao_quebrar(row)
        for i, cab in enumerate(cabecalhos):
            cel = row.cells[i]
            cel.width = Cm(larguras[i])
            _cel_fundo(cel, AZUL)
            _cel_margens(cel)
            escrever(cel, cab, True, "FFFFFF", "center" if i else "left")

    for idx, linha in enumerate(linhas):
        row = t.add_row()
        _nao_quebrar(row)
        ultima = linha_total and idx == len(linhas) - 1
        for i, valor in enumerate(linha):
            cel = row.cells[i]
            cel.width = Cm(larguras[i])
            _cel_margens(cel)
            if ultima:
                _cel_fundo(cel, FUNDO_AZUL)
            elif idx % 2 == 1:
                _cel_fundo(cel, FUNDO_CLARO)
            escrever(cel, valor, negrito=ultima, alinh=alinhamentos[i],
                     cor=AZUL if ultima else GRAFITE)

    _larguras(t, larguras)
    esp = doc.add_paragraph()
    esp.paragraph_format.space_after = Pt(0)
    esp.paragraph_format.space_before = Pt(0)
    esp.paragraph_format.line_spacing = 0.6


def tabela_kv(doc, pares, largura_rotulo=4.6, largura_valor=12.4) -> None:
    t = doc.add_table(rows=0, cols=2)
    t.autofit = False
    _tabela_bordas(t)
    for rotulo, valor in pares:
        if valor in (None, "", []):
            continue
        row = t.add_row()
        _nao_quebrar(row)
        c0, c1 = row.cells
        c0.width, c1.width = Cm(largura_rotulo), Cm(largura_valor)
        _cel_fundo(c0, FUNDO_CLARO)
        _cel_margens(c0), _cel_margens(c1)
        c0.text = ""
        p = c0.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(str(rotulo))
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor.from_string(AZUL)
        c1.text = ""
        p = c1.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(str(valor))
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor.from_string(GRAFITE)
    _larguras(t, [largura_rotulo, largura_valor])
    esp = doc.add_paragraph()
    esp.paragraph_format.space_after = Pt(0)
    esp.paragraph_format.space_before = Pt(0)
    esp.paragraph_format.line_spacing = 0.6


def assinaturas(doc, rt: dict, cliente: dict, local: str, data_txt: str) -> None:
    t = doc.add_table(rows=1, cols=2)
    _sem_bordas(t)
    t.autofit = False
    esq, dir_ = t.rows[0].cells
    esq.width, dir_.width = Cm(8.5), Cm(8.5)
    for cel, titulo, linhas in (
        (esq, "CONTRATADA", [rt.get("nome", ""), rt.get("titulo", ""),
                             f"CREA: {rt.get('crea','')}" if rt.get("crea") else ""]),
        (dir_, "CONTRATANTE", [cliente.get("nome", ""), cliente.get("documento", ""), ""]),
    ):
        cel.text = ""
        p = cel.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(26)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("_" * 34)
        r.font.size = Pt(9)
        for linha in linhas:
            if not linha:
                continue
            pp = cel.add_paragraph()
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pp.paragraph_format.space_after = Pt(0)
            rr = pp.add_run(linha)
            rr.font.size = Pt(9)
            rr.font.color.rgb = RGBColor.from_string(GRAFITE)
        pt = cel.add_paragraph()
        pt.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rr = pt.add_run(titulo)
        rr.bold = True
        rr.font.size = Pt(7.5)
        rr.font.color.rgb = RGBColor.from_string(CINZA)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    r = p.add_run(f"{local}, {data_txt}.")
    r.font.size = Pt(9.5)
    _larguras(t, [8.5, 8.5])


# ------------------------------------------------------------------ conteúdo
def gerar_docx(d: dict, calc: dict, destino: Path) -> Path:
    doc = Document()
    montar_folha(doc)

    p = d["proposta"]
    cli = d["cliente"]
    imv = d["imovel"]
    obj = d["objeto"]
    rt = d["responsavel_tecnico"]
    dt: date = p["data"]

    # ---- bloco de título
    pnum = doc.add_paragraph()
    pnum.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pnum.paragraph_format.space_before = Pt(2)
    pnum.paragraph_format.space_after = Pt(0)
    r = pnum.add_run(f"PROPOSTA Nº {p['numero']}" if p.get("numero") else "PROPOSTA")
    r.font.size = Pt(8.5)
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(CINZA)

    ptit = doc.add_paragraph()
    ptit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ptit.paragraph_format.space_before = Pt(2)
    ptit.paragraph_format.space_after = Pt(2)
    r = ptit.add_run(str(p["titulo"]).upper())
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor.from_string(AZUL)

    if p.get("subtitulo"):
        psub = doc.add_paragraph()
        psub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        psub.paragraph_format.space_after = Pt(10)
        r = psub.add_run(str(p["subtitulo"]))
        r.font.size = Pt(10.5)
        r.italic = True
        r.font.color.rgb = RGBColor.from_string(CINZA)
        _borda_paragrafo(psub, "bottom", BORDA, 6, 8)

    # ---- 1. identificação
    titulo_secao(doc, "1", "Identificação")
    local = " / ".join(x for x in [imv.get("municipio"), imv.get("uf")] if x)
    tabela_kv(doc, [
        ("Cliente", cli.get("nome")),
        ("CPF / CNPJ", cli.get("documento")),
        ("Endereço", cli.get("endereco")),
        ("Contato", " · ".join(x for x in [cli.get("contato"), cli.get("email")] if x)),
        ("Objeto", obj.get("resumo") or p.get("subtitulo")),
        ("Imóvel", imv.get("denominacao")),
        ("Localização", local),
        ("Data da proposta", C.data_longa(dt)),
        ("Validade", f"{p['validade_dias']} dias corridos"),
        ("Versão", f"v{int(p['versao']):02d}"),
    ])

    # ---- 2. apresentação
    if d.get("apresentacao"):
        titulo_secao(doc, "2", "Apresentação")
        for txt in d["apresentacao"]:
            par(doc, str(txt))

    # ---- 3. objeto
    titulo_secao(doc, "3", "Objeto")
    if obj.get("resumo"):
        par(doc, str(obj["resumo"]))
    if obj.get("finalidade"):
        par(doc, f"Finalidade: {obj['finalidade']}", italico=True, tamanho=10)
    if imv.get("area_ha") or imv.get("perimetro_m") or imv.get("matricula"):
        sub(doc, "3.1 Dados do imóvel")
        tabela_kv(doc, [
            ("Denominação", imv.get("denominacao")),
            ("Matrícula", imv.get("matricula")),
            ("Área", f"{C.num(imv['area_ha'], 4)} ha ({C.num(imv['area_ha'] * 10000, 2)} m²)"
             if imv.get("area_ha") else ""),
            ("Perímetro", f"{C.num(imv['perimetro_m'])} m" if imv.get("perimetro_m") else ""),
            ("Vértices", imv.get("vertices")),
            ("Município / UF", local),
        ])

    # ---- 4. escopo
    esc = d["escopo"]
    titulo_secao(doc, "4", "Escopo técnico dos serviços")
    etapas = esc.get("etapas") or []
    if etapas:
        sub(doc, "4.1 Etapas de execução")
        for i, et in enumerate(etapas, 1):
            p_et = doc.add_paragraph()
            p_et.paragraph_format.space_before = Pt(6)
            p_et.paragraph_format.space_after = Pt(2)
            p_et.paragraph_format.keep_with_next = True
            r = p_et.add_run(f"Etapa {i} — {et.get('titulo','')}")
            r.bold = True
            r.font.size = Pt(10)
            r.font.color.rgb = RGBColor.from_string(AZUL2)
            if et.get("descricao"):
                par(doc, str(et["descricao"]), espaco=4)
    if esc.get("entregaveis"):
        sub(doc, "4.2 Produtos e entregáveis")
        bullets(doc, esc["entregaveis"])
    if esc.get("nao_inclusos"):
        sub(doc, "4.3 Itens não inclusos")
        bullets(doc, esc["nao_inclusos"], marcador="–")

    # ---- 5. metodologia
    met = d["metodologia"]
    titulo_secao(doc, "5", "Metodologia e normas técnicas")
    for txt in met.get("paragrafos") or []:
        par(doc, str(txt))
    if met.get("normas"):
        sub(doc, "5.1 Referências normativas")
        bullets(doc, met["normas"], marcador="–")

    # ---- 6. cronograma
    cron = d["cronograma"]
    titulo_secao(doc, "6", "Cronograma de execução")
    linhas = [[str(e.get("fase", i + 1)), str(e.get("descricao", "")), str(e.get("prazo", ""))]
              for i, e in enumerate(cron.get("etapas") or [])]
    if linhas:
        tabela(doc, ["Fase", "Descrição", "Prazo"], linhas, [2.2, 11.0, 3.8],
               ["center", "left", "center"])
    if cron.get("prazo_dias"):
        par(doc, f"Prazo total estimado: {cron['prazo_dias']} dias corridos, contados da "
                 f"assinatura do contrato e da entrega da documentação prévia pelo contratante.",
            italico=True, tamanho=9.5)
    if cron.get("observacao"):
        par(doc, str(cron["observacao"]), tamanho=9.5)

    # ---- 7. investimento
    inv = d["investimento"]
    titulo_secao(doc, "7", "Investimento")

    def linha_item(it, total):
        return [str(it.get("item", "")), C.num(it.get("qtd", 1), 2), str(it.get("unidade", "un")),
                C.brl(it.get("valor_unitario") or it.get("valor") or 0), C.brl(total)]

    if calc["honorarios"]:
        sub(doc, "7.1 Honorários profissionais")
        linhas = [linha_item(it, v) for it, v in calc["honorarios"]]
        linhas.append(["Subtotal — honorários profissionais", "", "", "", C.brl(calc["soma_honorarios"])])
        tabela(doc, ["Item", "Qtd.", "Un.", "Valor unitário", "Total"],
               linhas, [7.3, 1.6, 2.0, 3.0, 3.1],
               ["left", "center", "center", "right", "right"], linha_total=True)

    if calc["despesas"]:
        sub(doc, "7.2 Despesas operacionais")
        linhas = [linha_item(it, v) for it, v in calc["despesas"]]
        linhas.append(["Subtotal — despesas operacionais", "", "", "", C.brl(calc["soma_despesas"])])
        tabela(doc, ["Item", "Qtd.", "Un.", "Valor unitário", "Total"],
               linhas, [7.3, 1.6, 2.0, 3.0, 3.1],
               ["left", "center", "center", "right", "right"], linha_total=True)

    if calc["tributos"]:
        sub(doc, "7.3 Tributos incidentes")
        linhas = []
        for it, v in calc["tributos"]:
            base_txt = (C.brl(it["base_calculo"]) if it.get("base_calculo")
                        else str(it.get("base") or "Honorários").replace("_", " ").capitalize())
            linhas.append([str(it.get("item", "")), base_txt,
                           C.pct(it.get("aliquota")) if it.get("aliquota") else "—", C.brl(v)])
        linhas.append(["Subtotal — tributos", "", "", C.brl(calc["soma_tributos"])])
        tabela(doc, ["Tributo", "Base de cálculo", "Alíquota", "Valor"],
               linhas, [7.4, 5.6, 2.0, 2.0],
               ["left", "left", "center", "right"], linha_total=True)

    if calc["terceiros"]:
        sub(doc, "7.4 Taxas e emolumentos de terceiros (informativos)")
        linhas = [[str(it.get("item", "")),
                   str(it.get("observacao") or "Responsabilidade do contratante"),
                   C.brl(v)] for it, v in calc["terceiros"]]
        linhas.append(["Subtotal informativo — não integra o valor da proposta", "", C.brl(calc["soma_terceiros"])])
        tabela(doc, ["Taxa / emolumento", "Observação", "Valor estimado"],
               linhas, [6.8, 6.6, 3.6],
               ["left", "left", "right"], linha_total=True)
        par(doc, "Os valores acima são estimativas informativas, recolhidas diretamente pelo "
                 "contratante aos órgãos competentes, e não estão incluídos no valor global desta proposta.",
            italico=True, tamanho=8.5, cor=CINZA)

    sub(doc, "7.5 Resumo do investimento")
    linhas = [["Honorários profissionais", C.brl(calc["soma_honorarios"])],
              ["Despesas operacionais", C.brl(calc["soma_despesas"])],
              ["Tributos incidentes", C.brl(calc["soma_tributos"])]]
    if calc["desconto"]:
        linhas.append([f"Desconto comercial ({C.pct(calc['desconto_pct'])})", f"- {C.brl(calc['desconto'])}"])
    linhas.append(["VALOR GLOBAL DA PROPOSTA", C.brl(calc["liquido"])])
    tabela(doc, ["Descrição", "Valor"], linhas, [12.0, 5.0], ["left", "right"], linha_total=True)

    par(doc, f"Valor global: {C.brl(calc['liquido'])} ({calc['extenso']}).", negrito=True, tamanho=10.5)
    if calc.get("por_hectare") and d["imovel"].get("area_ha"):
        par(doc, f"Equivalência: {C.brl(calc['por_hectare'])} por hectare sobre os honorários "
                 f"profissionais, considerada a área de {C.num(d['imovel']['area_ha'], 4)} ha.",
            tamanho=9, cor=CINZA)
    if inv.get("desconto", {}).get("justificativa"):
        par(doc, f"Condição do desconto: {inv['desconto']['justificativa']}", tamanho=9, cor=CINZA)

    # ---- 8. pagamento
    pag = d["pagamento"]
    titulo_secao(doc, "8", "Condições de pagamento")
    conds = pag.get("condicoes") or []
    if conds:
        linhas = [[str(c.get("etapa", "")), C.pct(c.get("percentual")), C.brl(c.get("valor"))]
                  for c in conds]
        tabela(doc, ["Etapa", "%", "Valor"], linhas, [11.0, 2.4, 3.6],
               ["left", "center", "right"])
    if pag.get("forma"):
        par(doc, str(pag["forma"]))
    if pag.get("observacao"):
        par(doc, str(pag["observacao"]), tamanho=9.5)

    # ---- 9. premissas
    if d.get("premissas"):
        titulo_secao(doc, "9", "Premissas, responsabilidades e exclusões")
        bullets(doc, d["premissas"], marcador="–")

    # ---- 10. validade
    titulo_secao(doc, "10", "Validade e disposições gerais")
    par(doc, f"Esta proposta é válida por {p['validade_dias']} dias corridos a contar de "
             f"{C.data_longa(dt)}, sujeita a reajuste após esse prazo em razão de variação de "
             f"custos, tributos ou condições de mercado.")
    for txt in d.get("disposicoes") or []:
        par(doc, str(txt))

    # ---- 11. responsabilidade técnica
    titulo_secao(doc, "11", "Responsabilidade técnica e assinatura")
    if rt.get("nome"):
        par(doc, f"Os serviços serão executados sob a responsabilidade técnica de "
                 f"{rt.get('nome')}, {rt.get('titulo','')}"
                 + (f", registrado no {rt.get('crea')}" if rt.get("crea") else "")
                 + ", com emissão da respectiva Anotação de Responsabilidade Técnica (ART) "
                   "junto ao CREA-PE, cujo recolhimento integra o valor desta proposta."
                 if rt.get("art_inclusa", True) else "")
    par(doc, "Aprovada a proposta, este documento passa a integrar o contrato de prestação de "
             "serviços, mediante assinatura das partes abaixo identificadas.")
    assinaturas(doc, rt, cli, rt.get("local_assinatura") or d["proposta"].get("local_data") or "",
                C.data_longa(dt))

    destino.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(destino))
    return destino


# ------------------------------------------------------------------ markdown
def _md_tabela(cabecalhos, linhas) -> str:
    out = ["| " + " | ".join(cabecalhos) + " |",
           "|" + "|".join(["---"] * len(cabecalhos)) + "|"]
    for linha in linhas:
        out.append("| " + " | ".join(str(c).replace("|", "\\|") for c in linha) + " |")
    return "\n".join(out)


def gerar_markdown(d: dict, calc: dict) -> str:
    p, cli, imv, obj, rt = (d["proposta"], d["cliente"], d["imovel"],
                            d["objeto"], d["responsavel_tecnico"])
    dt: date = p["data"]
    L = []
    meta = {
        "numero": p.get("numero", ""),
        "data": dt.isoformat(),
        "versao": int(p["versao"]),
        "cliente": cli.get("nome", ""),
        "objeto": obj.get("resumo") or p.get("subtitulo", ""),
        "servico": p.get("servico", ""),
        "municipio": " / ".join(x for x in [imv.get("municipio"), imv.get("uf")] if x),
        "area_ha": imv.get("area_ha") or 0,
        "honorarios": calc["soma_honorarios"],
        "valor": calc["liquido"],
        "status": p.get("status", "enviada"),
    }
    import json as _json
    L.append(f"<!-- meta: {_json.dumps(meta, ensure_ascii=False)} -->")
    L.append(f"<!-- Proposta gerada pela skill proposta-geotecnologias em "
             f"{date.today():%Y-%m-%d} -->")
    L.append("")
    L.append(f"# {p['titulo']}")
    L.append("")
    if p.get("subtitulo"):
        L.append(f"**{p['subtitulo']}**")
        L.append("")
    L.append(f"**R. Oliveira Consultoria** — Regularização de Imóveis · "
             f"Praça Pedro Coutinho, 56, Anexo Aderita Rosa de Oliveira, Loja 02, Centro, "
             f"Paudalho/PE, CEP 55.825-000 · (81) 98157-1227 · rodrigocabral@uol.com.br")
    L.append("")
    L.append(_md_tabela(["Campo", "Conteúdo"], [
        ["Proposta nº", p.get("numero") or "—"],
        ["Data", dt.strftime("%d/%m/%Y")],
        ["Versão", f"v{int(p['versao']):02d}"],
        ["Validade", f"{p['validade_dias']} dias"],
        ["Cliente", cli.get("nome") or "—"],
        ["Objeto", obj.get("resumo") or p.get("subtitulo") or "—"],
        ["Valor global", f"**{C.brl(calc['liquido'])}**"],
    ]))
    L.append("")
    L.append("---")
    L.append("")

    L.append("## 1. Identificação")
    L.append("")
    local = " / ".join(x for x in [imv.get("municipio"), imv.get("uf")] if x)
    pares = [("Cliente", cli.get("nome")), ("CPF/CNPJ", cli.get("documento")),
             ("Endereço", cli.get("endereco")),
             ("Contato", " · ".join(x for x in [cli.get("contato"), cli.get("email")] if x)),
             ("Imóvel", imv.get("denominacao")), ("Localização", local),
             ("Data da proposta", C.data_longa(dt)),
             ("Validade", f"{p['validade_dias']} dias corridos")]
    L.append(_md_tabela(["Campo", "Conteúdo"], [[k, v or "—"] for k, v in pares]))
    L.append("")

    if d.get("apresentacao"):
        L.append("## 2. Apresentação")
        L.append("")
        for txt in d["apresentacao"]:
            L.append(str(txt))
            L.append("")

    L.append("## 3. Objeto")
    L.append("")
    if obj.get("resumo"):
        L.append(str(obj["resumo"]))
        L.append("")
    if obj.get("finalidade"):
        L.append(f"**Finalidade:** {obj['finalidade']}")
        L.append("")
    if imv.get("area_ha") or imv.get("matricula") or imv.get("perimetro_m"):
        L.append("### 3.1 Dados do imóvel")
        L.append("")
        linhas = []
        if imv.get("denominacao"):
            linhas.append(["Denominação", imv["denominacao"]])
        if imv.get("matricula"):
            linhas.append(["Matrícula", imv["matricula"]])
        if imv.get("area_ha"):
            linhas.append(["Área", f"{C.num(imv['area_ha'], 4)} ha ({C.num(imv['area_ha']*10000, 2)} m²)"])
        if imv.get("perimetro_m"):
            linhas.append(["Perímetro", f"{C.num(imv['perimetro_m'])} m"])
        if imv.get("vertices"):
            linhas.append(["Vértices", imv["vertices"]])
        if local:
            linhas.append(["Município/UF", local])
        L.append(_md_tabela(["Item", "Valor"], linhas))
        L.append("")

    esc = d["escopo"]
    L.append("## 4. Escopo técnico dos serviços")
    L.append("")
    if esc.get("etapas"):
        L.append("### 4.1 Etapas de execução")
        L.append("")
        for i, et in enumerate(esc["etapas"], 1):
            L.append(f"**Etapa {i} — {et.get('titulo','')}**")
            L.append("")
            if et.get("descricao"):
                L.append(str(et["descricao"]))
                L.append("")
    if esc.get("entregaveis"):
        L.append("### 4.2 Produtos e entregáveis")
        L.append("")
        L.extend([f"- {x}" for x in esc["entregaveis"]])
        L.append("")
    if esc.get("nao_inclusos"):
        L.append("### 4.3 Itens não inclusos")
        L.append("")
        L.extend([f"- {x}" for x in esc["nao_inclusos"]])
        L.append("")

    met = d["metodologia"]
    L.append("## 5. Metodologia e normas técnicas")
    L.append("")
    for txt in met.get("paragrafos") or []:
        L.append(str(txt))
        L.append("")
    if met.get("normas"):
        L.append("### 5.1 Referências normativas")
        L.append("")
        L.extend([f"- {x}" for x in met["normas"]])
        L.append("")

    cron = d["cronograma"]
    L.append("## 6. Cronograma de execução")
    L.append("")
    if cron.get("etapas"):
        L.append(_md_tabela(["Fase", "Descrição", "Prazo"],
                            [[e.get("fase", i + 1), e.get("descricao", ""), e.get("prazo", "")]
                             for i, e in enumerate(cron["etapas"])]))
        L.append("")
    if cron.get("prazo_dias"):
        L.append(f"Prazo total estimado: **{cron['prazo_dias']} dias corridos**, contados da "
                 f"assinatura do contrato e da entrega da documentação prévia pelo contratante.")
        L.append("")

    L.append("## 7. Investimento")
    L.append("")
    if calc["honorarios"]:
        L.append("### 7.1 Honorários profissionais")
        L.append("")
        linhas = [[it.get("item", ""), C.num(it.get("qtd", 1)), it.get("unidade", "un"),
                   C.brl(it.get("valor_unitario") or it.get("valor") or 0), C.brl(v)]
                  for it, v in calc["honorarios"]]
        linhas.append(["**Subtotal**", "", "", "", f"**{C.brl(calc['soma_honorarios'])}**"])
        L.append(_md_tabela(["Item", "Qtd.", "Un.", "Valor unitário", "Total"], linhas))
        L.append("")
    if calc["despesas"]:
        L.append("### 7.2 Despesas operacionais")
        L.append("")
        linhas = [[it.get("item", ""), C.num(it.get("qtd", 1)), it.get("unidade", "un"),
                   C.brl(it.get("valor_unitario") or it.get("valor") or 0), C.brl(v)]
                  for it, v in calc["despesas"]]
        linhas.append(["**Subtotal**", "", "", "", f"**{C.brl(calc['soma_despesas'])}**"])
        L.append(_md_tabela(["Item", "Qtd.", "Un.", "Valor unitário", "Total"], linhas))
        L.append("")
    if calc["tributos"]:
        L.append("### 7.3 Tributos incidentes")
        L.append("")
        linhas = []
        for it, v in calc["tributos"]:
            base_txt = (C.brl(it["base_calculo"]) if it.get("base_calculo")
                        else str(it.get("base") or "honorarios").replace("_", " ").capitalize())
            linhas.append([it.get("item", ""), base_txt,
                           C.pct(it.get("aliquota")) if it.get("aliquota") else "—", C.brl(v)])
        linhas.append(["**Subtotal**", "", "", f"**{C.brl(calc['soma_tributos'])}**"])
        L.append(_md_tabela(["Tributo", "Base de cálculo", "Alíquota", "Valor"], linhas))
        L.append("")
    if calc["terceiros"]:
        L.append("### 7.4 Taxas e emolumentos de terceiros (informativos)")
        L.append("")
        linhas = [[it.get("item", ""), it.get("observacao") or "Responsabilidade do contratante",
                   C.brl(v)] for it, v in calc["terceiros"]]
        linhas.append(["**Subtotal informativo**", "não integra o valor da proposta",
                       f"**{C.brl(calc['soma_terceiros'])}**"])
        L.append(_md_tabela(["Taxa / emolumento", "Observação", "Valor estimado"], linhas))
        L.append("")
        L.append("> Valores estimativos, recolhidos diretamente pelo contratante aos órgãos "
                 "competentes; não estão incluídos no valor global desta proposta.")
        L.append("")

    L.append("### 7.5 Resumo do investimento")
    L.append("")
    linhas = [["Honorários profissionais", C.brl(calc["soma_honorarios"])],
              ["Despesas operacionais", C.brl(calc["soma_despesas"])],
              ["Tributos incidentes", C.brl(calc["soma_tributos"])]]
    if calc["desconto"]:
        linhas.append([f"Desconto comercial ({C.pct(calc['desconto_pct'])})", f"- {C.brl(calc['desconto'])}"])
    linhas.append(["**VALOR GLOBAL DA PROPOSTA**", f"**{C.brl(calc['liquido'])}**"])
    L.append(_md_tabela(["Descrição", "Valor"], linhas))
    L.append("")
    L.append(f"**Valor global: {C.brl(calc['liquido'])} ({calc['extenso']}).**")
    L.append("")
    if calc.get("por_hectare") and imv.get("area_ha"):
        L.append(f"*Equivalência: {C.brl(calc['por_hectare'])}/ha sobre honorários, "
                 f"para {C.num(imv['area_ha'], 4)} ha.*")
        L.append("")

    pag = d["pagamento"]
    L.append("## 8. Condições de pagamento")
    L.append("")
    if pag.get("condicoes"):
        L.append(_md_tabela(["Etapa", "%", "Valor"],
                            [[c.get("etapa", ""), C.pct(c.get("percentual")), C.brl(c.get("valor"))]
                             for c in pag["condicoes"]]))
        L.append("")
    if pag.get("forma"):
        L.append(str(pag["forma"]))
        L.append("")

    if d.get("premissas"):
        L.append("## 9. Premissas, responsabilidades e exclusões")
        L.append("")
        L.extend([f"- {x}" for x in d["premissas"]])
        L.append("")

    L.append("## 10. Validade e disposições gerais")
    L.append("")
    L.append(f"Esta proposta é válida por {p['validade_dias']} dias corridos a contar de "
             f"{C.data_longa(dt)}.")
    L.append("")
    for txt in d.get("disposicoes") or []:
        L.append(str(txt))
        L.append("")

    L.append("## 11. Responsabilidade técnica e assinatura")
    L.append("")
    if rt.get("nome"):
        L.append(f"Responsável técnico: **{rt.get('nome')}** — {rt.get('titulo','')}"
                 + (f" — CREA {rt.get('crea')}" if rt.get("crea") else ""))
        L.append("")
    L.append("| CONTRATADA | CONTRATANTE |")
    L.append("|---|---|")
    L.append(f"| {rt.get('nome','')} | {cli.get('nome','')} |")
    L.append(f"| {rt.get('titulo','')} | {cli.get('documento','')} |")
    L.append("")
    L.append(f"{rt.get('local_assinatura') or 'Recife/PE'}, {C.data_longa(dt)}.")
    L.append("")
    L.append("---")
    L.append("")
    L.append("*Documento gerado pela skill `proposta-geotecnologias` — cópia de arquivo em "
             "Markdown para versionamento e consulta.*")
    return "\n".join(L) + "\n"


# ------------------------------------------------------------------ CLI
def converter_pdf(docx_path: Path) -> Path | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("[aviso] LibreOffice não encontrado; PDF não gerado.", file=sys.stderr)
        return None
    out = docx_path.parent
    res = subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                          str(out), str(docx_path)],
                         capture_output=True, text=True, timeout=180)
    pdf = docx_path.with_suffix(".pdf")
    if res.returncode != 0 or not pdf.exists():
        print(f"[aviso] falha ao converter PDF: {res.stderr[:300]}", file=sys.stderr)
        return None
    return pdf


def main() -> None:
    ap = argparse.ArgumentParser(description="Gera proposta técnica em DOCX, PDF e Markdown.")
    ap.add_argument("dados", help="arquivo YAML com os dados da proposta")
    ap.add_argument("--saida", default="saida_proposta", help="diretório de saída")
    ap.add_argument("--sem-pdf", action="store_true", help="não converter para PDF")
    args = ap.parse_args()

    d = C.carregar_yaml(args.dados)
    calc = C.calcular(d)
    saida = Path(args.saida).expanduser().resolve()
    saida.mkdir(parents=True, exist_ok=True)

    base = f"Proposta_{d['proposta'].get('numero') or 'sem-numero'}_" \
           f"{C.slug(d['cliente'].get('nome','cliente'), 30)}_{d['proposta']['data']:%Y-%m-%d}_" \
           f"v{int(d['proposta']['versao']):02d}"
    docx_path = gerar_docx(d, calc, saida / f"{base}.docx")
    md_path = saida / C.nome_arquivo(d)
    md_path.write_text(gerar_markdown(d, calc), encoding="utf-8")

    pdf_path = None if args.sem_pdf else converter_pdf(docx_path)

    print(C.dump_json({
        "docx": str(docx_path),
        "pdf": str(pdf_path) if pdf_path else None,
        "md": str(md_path),
        "nome_md": md_path.name,
        "valor_global": calc["liquido"],
        "valor_global_fmt": C.brl(calc["liquido"]),
        "extenso": calc["extenso"],
        "honorarios": calc["soma_honorarios"],
        "despesas": calc["soma_despesas"],
        "tributos": calc["soma_tributos"],
        "terceiros_informativo": calc["soma_terceiros"],
        "desconto": calc["desconto"],
        "soma_percentuais_pagamento": calc["soma_percentuais"],
        "por_hectare": calc["por_hectare"],
    }))


if __name__ == "__main__":
    main()
