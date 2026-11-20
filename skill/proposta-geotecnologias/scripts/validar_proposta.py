#!/usr/bin/env python3
"""Valida o YAML da proposta antes de gerar/publicar o documento.

Uso:
    python3 validar_proposta.py dados.yaml [--md arquivo.md] [--docx arquivo.docx]

Retorna código 1 se houver erro impeditivo; avisos não bloqueiam.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _comum as C  # noqa: E402
C.garantir_dependencias()

PLACEHOLDERS = re.compile(r"(000\.000|0000000000|XXX|\[|\]|\{\{|\}\}|exemplo\.com)", re.I)


def validar(d: dict, calc: dict) -> tuple[list[str], list[str]]:
    erros: list[str] = []
    avisos: list[str] = []
    p, cli, obj, rt = d["proposta"], d["cliente"], d["objeto"], d["responsavel_tecnico"]

    # --- identificação
    if not cli.get("nome"):
        erros.append("cliente.nome é obrigatório")
    if not cli.get("documento"):
        avisos.append("cliente.documento (CPF/CNPJ) ausente — confirme antes de enviar")
    if not obj.get("resumo"):
        erros.append("objeto.resumo é obrigatório (descrição objetiva do serviço)")
    if not p.get("numero"):
        avisos.append("proposta.numero ausente — o controle interno e o nome dos arquivos perdem rastreabilidade")
    if not rt.get("nome"):
        erros.append("responsavel_tecnico.nome é obrigatório")
    if not rt.get("crea"):
        avisos.append("responsavel_tecnico.crea ausente")

    # --- valores
    if calc["soma_honorarios"] <= 0:
        erros.append("investimento.honorarios deve ter ao menos um item com valor > 0")
    if calc["liquido"] <= 0:
        erros.append("valor global calculado é zero — revise honorários/desconto")
    if calc["desconto_pct"] > 8:
        avisos.append(f"desconto de {calc['desconto_pct']:.2f}% acima do máximo padrão (8%) — exige aprovação")
    for it, v in calc["honorarios"] + calc["despesas"]:
        if v <= 0:
            avisos.append(f"item com total zero: {it.get('item','(sem descrição)')}")
        if not it.get("item"):
            erros.append("item de investimento sem descrição")

    # --- pagamento
    if not d["pagamento"].get("condicoes"):
        erros.append("pagamento.condicoes é obrigatório")
    elif abs(calc["soma_percentuais"] - 100) > 0.01:
        erros.append(f"percentuais de pagamento somam {calc['soma_percentuais']:.2f}% (deveriam somar 100%)")

    # --- cronograma e escopo
    if not d["cronograma"].get("prazo_dias"):
        avisos.append("cronograma.prazo_dias ausente — o prazo não aparecerá no documento")
    if not d["escopo"].get("entregaveis"):
        avisos.append("escopo.entregaveis vazio — liste os produtos para evitar questionamento")
    if not d["escopo"].get("nao_inclusos"):
        avisos.append("escopo.nao_inclusos vazio — recomenda-se explicitar exclusões")
    if not d["metodologia"].get("normas"):
        avisos.append("metodologia.normas vazio — cite a norma técnica aplicável")

    # --- área x serviço por hectare
    area = d["imovel"].get("area_ha") or 0
    por_ha = any(str(it.get("unidade", "")).lower() in ("ha", "hectare", "hectares")
                 for it in d["investimento"].get("honorarios", []))
    if por_ha and area <= 0:
        erros.append("há item cobrado por hectare, mas imovel.area_ha está vazio")
    if area > 0 and calc.get("por_hectare"):
        avisos.append(f"equivalência de honorários: {C.brl(calc['por_hectare'])}/ha")

    # --- terceiros
    if calc["soma_terceiros"] > 0:
        avisos.append("taxas de terceiros são informativas e não entram no valor global — confirme a redação")

    # --- datas
    dt: date = p["data"]
    if dt > date.today():
        avisos.append(f"data da proposta é futura ({dt:%d/%m/%Y})")
    if (date.today() - dt).days > 180:
        avisos.append(f"proposta com data antiga ({(date.today() - dt).days} dias) — verifique validade")

    # --- placeholders
    texto = json.dumps(d, ensure_ascii=False, default=str)
    if PLACEHOLDERS.search(texto):
        avisos.append("há marcadores de exemplo/placeholder nos dados (CPF 000..., CREA 000..., [texto])")

    return erros, avisos


def validar_arquivos(md: Path | None, docx: Path | None, calc: dict) -> list[str]:
    avisos = []
    if md:
        if not md.exists():
            avisos.append(f"Markdown não encontrado: {md}")
        else:
            txt = md.read_text(encoding="utf-8")
            esperado = C.brl(calc["liquido"])
            if esperado not in txt:
                avisos.append(f"o valor global {esperado} não aparece no Markdown")
            if calc["extenso"] not in txt:
                avisos.append("valor por extenso não encontrado no Markdown")
            if "<!-- meta:" not in txt:
                avisos.append("bloco de metadados ausente no Markdown (índice do repositório ficará incompleto)")
    if docx and not docx.exists():
        avisos.append(f"DOCX não encontrado: {docx}")
    return avisos


def main() -> None:
    ap = argparse.ArgumentParser(description="Valida os dados da proposta.")
    ap.add_argument("dados")
    ap.add_argument("--md")
    ap.add_argument("--docx")
    args = ap.parse_args()

    d = C.carregar_yaml(args.dados)
    calc = C.calcular(d)
    erros, avisos = validar(d, calc)
    avisos += validar_arquivos(Path(args.md) if args.md else None,
                               Path(args.docx) if args.docx else None, calc)

    print(json.dumps({
        "arquivo": str(args.dados),
        "nome_md": C.nome_arquivo(d),
        "valor_global": C.brl(calc["liquido"]),
        "extenso": calc["extenso"],
        "erros": erros,
        "avisos": avisos,
        "ok": not erros,
    }, ensure_ascii=False, indent=2))
    sys.exit(1 if erros else 0)


if __name__ == "__main__":
    main()
