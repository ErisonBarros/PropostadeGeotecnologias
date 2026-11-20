#!/usr/bin/env python3
"""Gera a planilha de despesas de execução em .xls (Excel 97-2003).

Uso:
    python3 gerar_despesas.py dados.yaml [--saida DIR]

Saída: despesa_AAAA_MM_DD_vNN.xls, com três blocos (mão de obra, material e serviços
terceirizados), fórmulas vivas (quantidade × valor), resumo e conciliação com a proposta.
"""

from __future__ import annotations

import argparse
import math
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _comum as C  # noqa: E402

C.garantir_dependencias()

import xlwt  # noqa: E402
from xlwt import Formula, Workbook  # noqa: E402

MOEDA = '#,##0.00'
PCT = '0.0%'

# índices 0x21..0x24 da paleta, redefinidos com as cores da marca
COR_AZUL = 0x21
COR_AZUL2 = 0x22
COR_CINZA = 0x23
COR_AZUL_CLARO = 0x24
COR_CINZA_TXT = 0x25
BRANCO = 0x01
GRAFITE = 0x08

LARGURAS = [46, 16, 11, 16, 26, 16]  # em caracteres (colunas A..F)
_E: dict = {}


# ------------------------------------------------------------------ estilos
def _fonte(bold=False, italic=False, altura=200, cor=GRAFITE):
    f = xlwt.Font()
    f.name = "Calibri"
    f.bold = bold
    f.italic = italic
    f.height = altura
    f.colour_index = cor
    return f


def _borda(estilo):
    estilo.borders = xlwt.Borders()
    for lado in ("left", "right", "top", "bottom"):
        setattr(estilo.borders, lado, xlwt.Borders.THIN)
    estilo.borders.left_colour = COR_AZUL2
    estilo.borders.right_colour = COR_AZUL2
    estilo.borders.top_colour = COR_AZUL2
    estilo.borders.bottom_colour = COR_AZUL2
    return estilo


def _fundo(estilo, cor):
    estilo.pattern = xlwt.Pattern()
    estilo.pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    estilo.pattern.pattern_fore_colour = cor
    return estilo


def _estilos() -> dict:
    e = {}

    s = xlwt.XFStyle()
    s.font = _fonte(bold=True, altura=320, cor=COR_AZUL)
    s.alignment = xlwt.Alignment()
    s.alignment.horz = xlwt.Alignment.HORZ_CENTER
    s.alignment.vert = xlwt.Alignment.VERT_CENTER
    e["titulo"] = s

    s = xlwt.XFStyle()
    s.font = _fonte(italic=True, altura=190, cor=COR_CINZA_TXT)
    s.alignment = xlwt.Alignment()
    s.alignment.horz = xlwt.Alignment.HORZ_CENTER
    s.alignment.vert = xlwt.Alignment.VERT_CENTER
    e["subtitulo"] = s

    s = xlwt.XFStyle()
    s.font = _fonte(bold=True, altura=200, cor=COR_AZUL)
    s.alignment = xlwt.Alignment()
    s.alignment.vert = xlwt.Alignment.VERT_CENTER
    s.alignment.wrap = 1
    s.alignment.indent = 1
    s.alignment.horz = xlwt.Alignment.HORZ_LEFT
    e["rotulo"] = s

    s = xlwt.XFStyle()
    s.font = _fonte(altura=200)
    s.alignment = xlwt.Alignment()
    s.alignment.vert = xlwt.Alignment.VERT_CENTER
    s.alignment.wrap = 1
    s.alignment.indent = 1
    s.alignment.horz = xlwt.Alignment.HORZ_LEFT
    e["valor"] = s

    s = _fundo(xlwt.XFStyle(), COR_AZUL_CLARO)
    s.font = _fonte(bold=True, altura=210, cor=COR_AZUL)
    s.alignment = xlwt.Alignment()
    s.alignment.vert = xlwt.Alignment.VERT_CENTER
    e["secao"] = s

    s = _fundo(_borda(xlwt.XFStyle()), COR_AZUL)
    s.font = _fonte(bold=True, altura=190, cor=BRANCO)
    s.alignment = xlwt.Alignment()
    s.alignment.horz = xlwt.Alignment.HORZ_CENTER
    s.alignment.vert = xlwt.Alignment.VERT_CENTER
    s.alignment.wrap = 1
    e["cabecalho"] = s

    def celula(moeda=False, negrito=False, fundo=None, centro=False, cor=GRAFITE):
        s = _borda(xlwt.XFStyle())
        s.font = _fonte(bold=negrito, altura=195, cor=cor)
        if moeda:
            s.num_format_str = MOEDA
        if fundo:
            _fundo(s, fundo)
        s.alignment = xlwt.Alignment()
        s.alignment.vert = xlwt.Alignment.VERT_CENTER
        s.alignment.wrap = 1
        s.alignment.indent = 1
        s.alignment.horz = (xlwt.Alignment.HORZ_CENTER if centro
                            else xlwt.Alignment.HORZ_LEFT)
        return s

    e["txt"] = celula()
    e["txt_c"] = celula(centro=True)
    e["num"] = celula(moeda=True)
    e["pct"] = celula(centro=True)
    e["pct"].num_format_str = PCT
    e["total_txt"] = celula(negrito=True, fundo=COR_AZUL_CLARO, cor=COR_AZUL)
    e["total"] = celula(moeda=True, negrito=True, fundo=COR_AZUL_CLARO, cor=COR_AZUL)
    e["resumo_rot"] = celula(negrito=True)
    e["resumo_val"] = celula(moeda=True, negrito=True)
    e["destaque"] = celula(moeda=True, negrito=True, fundo=COR_AZUL_CLARO, cor=COR_AZUL)
    e["destaque_txt"] = celula(negrito=True, fundo=COR_AZUL_CLARO, cor=COR_AZUL)

    s = xlwt.XFStyle()
    s.font = _fonte(italic=True, altura=180, cor=COR_CINZA_TXT)
    s.alignment = xlwt.Alignment()
    s.alignment.wrap = 1
    s.alignment.vert = xlwt.Alignment.VERT_TOP
    e["obs"] = s
    return e


# ------------------------------------------------------------------ layout
def _altura(textos_larguras, minimo=340, por_linha=330):
    linhas = 1
    for texto, largura in textos_larguras:
        if not texto:
            continue
        texto = str(texto)
        estimado = math.ceil(len(texto) / max(int(largura * 0.85) - 1, 4))
        linhas = max(linhas, estimado)
    return max(minimo, linhas * por_linha + 80)


def _barra(ws, linha, texto, ncols=6):
    ws.write_merge(linha, linha, 0, ncols - 1, texto, _E["secao"])
    ws.row(linha).height = 460
    return linha + 1


def _cabecalho(ws, linha, nomes):
    for i, nome in enumerate(nomes):
        ws.write(linha, i, nome, _E["cabecalho"])
    ws.row(linha).height = _altura([(n, LARGURAS[i]) for i, n in enumerate(nomes)],
                                   minimo=560, por_linha=250)
    return linha + 1


def _ajustar(ws, linha, pares, minimo=300):
    ws.row(linha).height = _altura(pares, minimo=minimo)


def gerar_xls(d: dict, des: dict, destino: Path) -> Path:
    global _E
    wb = Workbook()
    wb.set_colour_RGB(COR_AZUL, 0x05, 0x0A, 0x30)
    wb.set_colour_RGB(COR_AZUL2, 0x03, 0x40, 0x6D)
    wb.set_colour_RGB(COR_CINZA, 0xF2, 0xF4, 0xF7)
    wb.set_colour_RGB(COR_AZUL_CLARO, 0xDC, 0xE6, 0xF1)
    wb.set_colour_RGB(COR_CINZA_TXT, 0x5A, 0x64, 0x72)
    _E = _estilos()

    ws = wb.add_sheet("Despesas")
    ws.set_portrait(False)
    ws.set_fit_num_pages(1)
    ws.set_fit_width_to_pages(1)
    ws.set_paper_size_code(9)  # A4
    ws.left_margin = 0.4
    ws.right_margin = 0.4
    ws.top_margin = 0.5
    ws.bottom_margin = 0.5
    for i, largura in enumerate(LARGURAS):
        ws.col(i).width = int(largura * 256)

    p, cli, imv, obj = d["proposta"], d["cliente"], d["imovel"], d["objeto"]
    dt: date = p["data"]
    rt = d["responsavel_tecnico"]

    ws.write_merge(0, 0, 0, 5, "PLANILHA DE DESPESAS DE EXECUÇÃO", _E["titulo"])
    ws.row(0).height = 640
    ws.write_merge(1, 1, 0, 5,
                   f"Proposta nº {p.get('numero') or '—'}  ·  {dt:%d/%m/%Y}  ·  "
                   f"versão v{int(p['versao']):02d}", _E["subtitulo"])
    ws.row(1).height = 320

    # --- identificação
    linha = 3
    identificacao = [
        ("Cliente", cli.get("nome") or "—"),
        ("Objeto", obj.get("resumo") or p.get("subtitulo") or "—"),
        ("Imóvel / local", " · ".join(x for x in [
            imv.get("denominacao"),
            "/".join(y for y in [imv.get("municipio"), imv.get("uf")] if y)] if x)),
        ("Área", f"{C.num(imv['area_ha'], 4)} ha" if imv.get("area_ha") else "—"),
        ("Responsável técnico", f"{rt.get('nome','—')} — {rt.get('titulo','')}".strip(" —")),
        ("Encargos sociais", f"{C.num(des['encargos_sociais'])}% sobre a mão de obra"),
    ]
    for rot, val in identificacao:
        ws.write(linha, 0, rot, _E["rotulo"])
        ws.write_merge(linha, linha, 1, 5, val, _E["valor"])
        _ajustar(ws, linha, [(val, sum(LARGURAS[1:])), (rot, LARGURAS[0])])
        linha += 1

    # --- A. mão de obra
    linha += 1
    linha = _barra(ws, linha, "A. MÃO DE OBRA")
    linha = _cabecalho(ws, linha, [
        "Função / cargo", "Qtd. profissionais", "Diárias", "Valor da diária (R$)",
        "Encargos (%)", "Total (R$)"])
    inicio = linha
    for it in des["mao_de_obra"]:
        ws.write(linha, 0, it.get("funcao", ""), _E["txt"])
        ws.write(linha, 1, it["qtd"], _E["txt_c"])
        ws.write(linha, 2, it["dias"], _E["txt_c"])
        ws.write(linha, 3, it["valor_diaria"], _E["num"])
        ws.write(linha, 4, des["encargos_sociais"] / 100.0, _E["pct"])
        ws.write(linha, 5, Formula(f"B{linha+1}*C{linha+1}*D{linha+1}*(1+E{linha+1})"),
                 _E["num"])
        _ajustar(ws, linha, [(it.get("funcao", ""), LARGURAS[0])])
        linha += 1
    if not des["mao_de_obra"]:
        ws.write_merge(linha, linha, 0, 5, "Sem itens de mão de obra informados.", _E["txt"])
        linha += 1
    fim = linha - 1
    ws.write(linha, 0, "Subtotal — mão de obra", _E["total_txt"])
    ws.write_merge(linha, linha, 1, 4, "", _E["total_txt"])
    ws.write(linha, 5, Formula(f"SUM(F{inicio+1}:F{fim+1})") if fim >= inicio else 0,
             _E["total"])
    ws.row(linha).height = 360
    sub_mao = linha
    linha += 2

    # --- B. material
    linha = _barra(ws, linha, "B. MATERIAL E INSUMOS")
    linha = _cabecalho(ws, linha, [
        "Item / descrição", "Quantidade", "Unidade", "Valor unitário (R$)",
        "Observação", "Total (R$)"])
    inicio = linha
    for it in des["materiais"]:
        ws.write(linha, 0, it.get("item", ""), _E["txt"])
        ws.write(linha, 1, it["qtd"], _E["txt_c"])
        ws.write(linha, 2, it.get("unidade", "un"), _E["txt_c"])
        ws.write(linha, 3, it["valor_unitario"], _E["num"])
        ws.write(linha, 4, it.get("observacao", ""), _E["txt"])
        ws.write(linha, 5, Formula(f"B{linha+1}*D{linha+1}"), _E["num"])
        _ajustar(ws, linha, [(it.get("item", ""), LARGURAS[0]),
                             (it.get("observacao", ""), LARGURAS[4])])
        linha += 1
    if not des["materiais"]:
        ws.write_merge(linha, linha, 0, 5, "Sem materiais informados.", _E["txt"])
        linha += 1
    fim = linha - 1
    ws.write(linha, 0, "Subtotal — material", _E["total_txt"])
    ws.write_merge(linha, linha, 1, 4, "", _E["total_txt"])
    ws.write(linha, 5, Formula(f"SUM(F{inicio+1}:F{fim+1})") if fim >= inicio else 0,
             _E["total"])
    ws.row(linha).height = 360
    sub_mat = linha
    linha += 2

    # --- C. serviços terceirizados
    linha = _barra(ws, linha, "C. SERVIÇOS TERCEIRIZADOS")
    linha = _cabecalho(ws, linha, [
        "Serviço / descrição", "Fornecedor", "Quantidade", "Unidade",
        "Valor unitário (R$)", "Total (R$)"])
    inicio = linha
    for it in des["servicos_terceirizados"]:
        ws.write(linha, 0, it.get("item", ""), _E["txt"])
        ws.write(linha, 1, it.get("fornecedor", "a definir"), _E["txt"])
        ws.write(linha, 2, it["qtd"], _E["txt_c"])
        ws.write(linha, 3, it.get("unidade", "serviço"), _E["txt_c"])
        ws.write(linha, 4, it["valor_unitario"], _E["num"])
        ws.write(linha, 5, Formula(f"C{linha+1}*E{linha+1}"), _E["num"])
        _ajustar(ws, linha, [(it.get("item", ""), LARGURAS[0]),
                             (it.get("fornecedor", ""), LARGURAS[1])])
        linha += 1
    if not des["servicos_terceirizados"]:
        ws.write_merge(linha, linha, 0, 5, "Sem serviços terceirizados informados.", _E["txt"])
        linha += 1
    fim = linha - 1
    ws.write(linha, 0, "Subtotal — serviços terceirizados", _E["total_txt"])
    ws.write_merge(linha, linha, 1, 4, "", _E["total_txt"])
    ws.write(linha, 5, Formula(f"SUM(F{inicio+1}:F{fim+1})") if fim >= inicio else 0,
             _E["total"])
    ws.row(linha).height = 360
    sub_ter = linha
    linha += 2

    # --- resumo
    linha = _barra(ws, linha, "RESUMO DAS DESPESAS DE EXECUÇÃO")
    primeira = linha
    resumo = [("Mão de obra", f"F{sub_mao+1}"),
              ("Material e insumos", f"F{sub_mat+1}"),
              ("Serviços terceirizados", f"F{sub_ter+1}")]
    for rot, ref in resumo:
        ws.write(linha, 0, rot, _E["resumo_rot"])
        ws.write(linha, 1, Formula(ref), _E["resumo_val"])
        ws.write(linha, 2, Formula(f"B{linha+1}/B{primeira+4}"), _E["pct"])
        ws.write_merge(linha, linha, 3, 5, "  participação no total", _E["txt"])
        ws.row(linha).height = 330
        linha += 1
    ws.write(linha, 0, "TOTAL GERAL DAS DESPESAS", _E["destaque_txt"])
    ws.write(linha, 1, Formula(f"SUM(B{primeira+1}:B{linha})"), _E["destaque"])
    ws.write_merge(linha, linha, 2, 5, "", _E["destaque_txt"])
    ws.row(linha).height = 400
    total_geral = linha
    linha += 2

    # --- conciliação
    linha = _barra(ws, linha, "CONCILIAÇÃO COM A PROPOSTA")
    itens_conc = [
        ("Despesas operacionais na proposta (item 7.2)", des["despesas_proposta"]),
        ("Material + serviços terceirizados desta planilha",
         Formula(f"F{sub_mat+1}+F{sub_ter+1}")),
    ]
    for rot, val in itens_conc:
        ws.write(linha, 0, rot, _E["resumo_rot"])
        ws.write(linha, 1, val, _E["resumo_val"])
        ws.write_merge(linha, linha, 2, 5, "", _E["txt"])
        _ajustar(ws, linha, [(rot, LARGURAS[0])], minimo=320)
        linha += 1
    ws.write(linha, 0, "Diferença (proposta − planilha)", _E["resumo_rot"])
    ws.write(linha, 1, Formula(f"B{linha-1}-B{linha}"), _E["resumo_val"])
    ws.write_merge(linha, linha, 2, 5, "  positivo = planilha abaixo do previsto", _E["txt"])
    _ajustar(ws, linha, [("Diferença (proposta − planilha)", LARGURAS[0])], minimo=320)
    linha += 1
    ws.write(linha, 0, "Mão de obra (embutida nos honorários)", _E["resumo_rot"])
    ws.write(linha, 1, Formula(f"F{sub_mao+1}"), _E["resumo_val"])
    ws.write_merge(linha, linha, 2, 5, "  custo direto, não lançado como despesa", _E["txt"])
    _ajustar(ws, linha, [("Mão de obra (embutida nos honorários)", LARGURAS[0])], minimo=320)
    linha += 1
    if imv.get("area_ha"):
        ws.write(linha, 0, "Custo total de execução por hectare", _E["resumo_rot"])
        ws.write(linha, 1, Formula(f"B{total_geral+1}/{C.num(imv['area_ha'], 4)}".replace(",", ".")),
                 _E["resumo_val"])
        ws.write_merge(linha, linha, 2, 5, "  R$/ha", _E["txt"])
        linha += 1

    # --- observações
    linha += 1
    ws.write(linha, 0, "Observações", _E["rotulo"])
    linha += 1
    obs = (d.get("custos_execucao", {}).get("observacoes")
           or "Custo direto de execução, sem margem, tributos ou BDI. As fórmulas permanecem "
              "vivas: altere quantidade, diária ou valor unitário e o total é recalculado.")
    ws.write_merge(linha, linha + 2, 0, 5, obs, _E["obs"])
    for r in range(linha, linha + 3):
        ws.row(r).height = 300

    # --- segunda aba: premissas
    ws2 = wb.add_sheet("Premissas")
    ws2.set_portrait(False)
    ws2.set_fit_num_pages(1)
    ws2.set_fit_width_to_pages(1)
    ws2.set_paper_size_code(9)
    ws2.left_margin = 0.4
    ws2.right_margin = 0.4
    ws2.top_margin = 0.5
    ws2.bottom_margin = 0.5
    ws2.col(0).width = int(26 * 256)
    ws2.col(1).width = int(78 * 256)
    ws2.write_merge(0, 0, 0, 1, "PREMISSAS E CRITÉRIOS DA PLANILHA", _E["titulo"])
    ws2.row(0).height = 560
    premissas = [
        ("Encargos sociais",
         f"{C.num(des['encargos_sociais'])}% aplicados sobre a mão de obra (INSS, FGTS, férias, "
         f"13º e demais incidências conforme o regime da empresa)."),
        ("Diárias",
         "Incluem hospedagem e alimentação quando a equipe pernoita no local; o deslocamento "
         "deve ser lançado em material, serviço terceirizado ou item próprio."),
        ("Material",
         "Insumos consumidos na execução (marcos, estacas, tinta, EPI, plotagem). Equipamentos "
         "próprios não são depreciados nesta planilha."),
        ("Serviços terceirizados",
         "Consultoria, laboratório, transporte, plotagem externa e demais serviços contratados "
         "de terceiros."),
        ("Não incluso",
         "Tributos sobre a receita (ISSQN), margem de lucro, BDI, taxas e emolumentos pagos "
         "pelo cliente."),
        ("Uso",
         "Planilha de custo interno. Não enviar ao cliente sem revisão do responsável técnico."),
    ]
    r = 2
    for rot, txt in premissas:
        ws2.write(r, 0, rot, _E["rotulo"])
        ws2.write(r, 1, txt, _E["obs"])
        ws2.row(r).height = _altura([(txt, 78)], minimo=420)
        r += 1

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(destino))
    return destino


def main() -> None:
    ap = argparse.ArgumentParser(description="Gera a planilha de despesas de execução (.xls).")
    ap.add_argument("dados")
    ap.add_argument("--saida", default="saida_proposta")
    args = ap.parse_args()

    d = C.carregar_yaml(args.dados)
    des = C.calcular_despesas(d)
    saida = Path(args.saida).expanduser().resolve()
    destino = saida / C.nomes_arquivos(d)["despesas"]
    gerar_xls(d, des, destino)
    print(C.dump_json({
        "xls": str(destino),
        "soma_mao_de_obra": des["soma_mao_de_obra"],
        "soma_materiais": des["soma_materiais"],
        "soma_terceiros": des["soma_terceiros"],
        "total": des["total"],
        "total_fmt": C.brl(des["total"]),
        "despesas_proposta": des["despesas_proposta"],
        "diferenca": des["diferenca"],
    }))


if __name__ == "__main__":
    main()
