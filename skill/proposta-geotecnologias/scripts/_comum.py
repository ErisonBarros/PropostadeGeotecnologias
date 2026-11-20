"""Utilitários compartilhados da skill proposta-geotecnologias.

Responsabilidades:
  * bootstrap de dependências (python-docx / PyYAML) em venv isolado;
  * formatação pt-BR (moeda, números, percentuais);
  * valor por extenso em reais;
  * leitura do YAML da proposta e normalização de campos;
  * cálculo da planilha de investimento (fonte única de verdade).
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
TEMPLATES = SKILL_DIR / "templates"
IDENTIDADE = TEMPLATES / "identidade"
CACHE = Path.home() / ".cache" / "proposta-geo"

# ---------------------------------------------------------------- dependências


def garantir_dependencias() -> None:
    """Garante python-docx e PyYAML; recria a si mesmo dentro de um venv."""
    faltando = []
    for mod in ("docx", "yaml"):
        try:
            __import__(mod)
        except ImportError:
            faltando.append(mod)
    if not faltando:
        return

    venv = CACHE / "venv"
    py = venv / "bin" / "python"
    if not py.exists():
        CACHE.mkdir(parents=True, exist_ok=True)
        print(f"[deps] criando venv em {venv}", file=sys.stderr)
        subprocess.run([sys.executable, "-m", "venv", str(venv)], check=True)
        subprocess.run(
            [str(py), "-m", "pip", "install", "--quiet", "--upgrade", "pip"],
            check=False,
        )
        subprocess.run(
            [str(py), "-m", "pip", "install", "--quiet", "python-docx", "PyYAML"],
            check=True,
        )
    # o python do venv é symlink para o interpretador do sistema: comparar prefix
    if Path(sys.prefix).resolve() != venv.resolve():
        script = Path(sys.argv[0]).resolve()
        if not script.exists():
            raise SystemExit(
                "python-docx/PyYAML ausentes e não foi possível reiniciar em venv. "
                f"Instale manualmente: {py} -m pip install python-docx PyYAML"
            )
        os.execv(str(py), [str(py), str(script)] + sys.argv[1:])


# ---------------------------------------------------------------- formatação

def brl(valor: float | int | None) -> str:
    """1234.5 -> 'R$ 1.234,50'."""
    v = float(valor or 0)
    s = f"{v:,.2f}"
    s = s.replace(",", "\x00").replace(".", ",").replace("\x00", ".")
    return f"R$ {s}"


def num(valor: float | int | None, casas: int = 2) -> str:
    v = float(valor or 0)
    s = f"{v:,.{casas}f}"
    return s.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def pct(valor: float | int | None, casas: int = 2) -> str:
    return f"{num(valor, casas)}%"


# ---------------------------------------------------------------- extenso

_UNI = ["zero", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove",
        "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete",
        "dezoito", "dezenove"]
_DEZ = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta",
        "oitenta", "noventa"]
_CEM = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos",
        "setecentos", "oitocentos", "novecentos"]


def _ate_999(n: int) -> str:
    if n == 0:
        return ""
    if n == 100:
        return "cem"
    if n < 20:
        return _UNI[n]
    if n < 100:
        d, u = divmod(n, 10)
        return _DEZ[d] + (" e " + _UNI[u] if u else "")
    c, r = divmod(n, 100)
    base = _CEM[c]
    if r == 0:
        return base
    return base + " e " + _ate_999(r)


def _inteiro_extenso(n: int) -> str:
    if n == 0:
        return "zero"
    blocos = []
    escalas = [(1_000_000_000, "bilhão", "bilhões"),
               (1_000_000, "milhão", "milhões"),
               (1_000, "mil", "mil")]
    resto = n
    for valor, sing, plur in escalas:
        q, resto = divmod(resto, valor)
        if q:
            if valor == 1_000 and q == 1:
                blocos.append("mil")
            else:
                nome = sing if q == 1 else plur
                blocos.append(f"{_ate_999(q)} {nome}".strip())
    if resto:
        blocos.append(_ate_999(resto))
    if len(blocos) == 1:
        return blocos[0]
    return ", ".join(blocos[:-1]) + " e " + blocos[-1]


def extenso(valor: float | int | None) -> str:
    """Valor monetário por extenso, ex.: 'dois mil e quinhentos reais e trinta centavos'."""
    total = round(float(valor or 0) * 100)
    reais, centavos = divmod(abs(total), 100)
    partes = []
    if reais:
        partes.append(f"{_inteiro_extenso(reais)} {'real' if reais == 1 else 'reais'}")
    if centavos:
        partes.append(f"{_inteiro_extenso(centavos)} {'centavo' if centavos == 1 else 'centavos'}")
    if not partes:
        return "zero real"
    txt = " e ".join(partes)
    return ("menos " if total < 0 else "") + txt


# ---------------------------------------------------------------- utilidades

def slug(texto: str, limite: int = 60) -> str:
    s = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Za-z0-9]+", "-", s).strip("-").lower()
    return s[:limite].strip("-") or "proposta"


def data_iso(valor: str | date | None) -> date:
    if isinstance(valor, date):
        return valor
    if not valor:
        return date.today()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(str(valor).strip(), fmt).date()
        except ValueError:
            continue
    raise SystemExit(f"Data inválida: {valor!r} (use AAAA-MM-DD)")


MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]


def data_longa(d: date) -> str:
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def carregar_yaml(caminho: str | Path) -> dict:
    import yaml

    p = Path(caminho)
    if not p.exists():
        raise SystemExit(f"Arquivo de dados não encontrado: {p}")
    with p.open("r", encoding="utf-8") as fh:
        dados = yaml.safe_load(fh) or {}
    if not isinstance(dados, dict):
        raise SystemExit("O YAML deve ser um mapa de chaves/valores.")
    return normalizar(dados)


def normalizar(d: dict) -> dict:
    """Preenche defaults e converte tipos numéricos com tolerância a texto pt-BR."""
    def fnum(v, padrao=0.0):
        if v is None or v == "":
            return padrao
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).replace("R$", "").replace(" ", "").strip()
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return padrao

    p = d.setdefault("proposta", {}) or {}
    p.setdefault("numero", "")
    p["data"] = data_iso(p.get("data"))
    p["versao"] = int(p.get("versao") or 1)
    p.setdefault("validade_dias", 30)
    p.setdefault("titulo", "PROPOSTA TÉCNICA E COMERCIAL")
    p.setdefault("subtitulo", "")
    p.setdefault("local_data", "")

    d.setdefault("cliente", {})
    d.setdefault("imovel", {})
    d.setdefault("objeto", {})
    d.setdefault("escopo", {})
    d.setdefault("metodologia", {})
    d.setdefault("cronograma", {})
    d.setdefault("pagamento", {})
    d.setdefault("responsavel_tecnico", {})
    d.setdefault("premissas", [])

    inv = d.setdefault("investimento", {})
    for chave in ("honorarios", "despesas", "tributos", "terceiros"):
        itens = inv.get(chave) or []
        for it in itens:
            for campo in ("qtd", "valor_unitario", "valor", "aliquota", "base_calculo"):
                if campo in it:
                    it[campo] = fnum(it[campo])
        inv[chave] = itens
    desc = inv.setdefault("desconto", {}) or {}
    desc["percentual"] = fnum(desc.get("percentual"))
    inv["desconto"] = desc

    d["imovel"]["area_ha"] = fnum(d["imovel"].get("area_ha"))
    d["imovel"]["perimetro_m"] = fnum(d["imovel"].get("perimetro_m"))
    return d


# ---------------------------------------------------------------- cálculos

def calcular(d: dict) -> dict:
    """Calcula a planilha de investimento. Fonte única para DOCX, PDF e MD."""
    inv = d["investimento"]

    def total_item(it):
        if it.get("valor") not in (None, "") and not it.get("qtd"):
            return float(it["valor"])
        qtd = float(it.get("qtd") or 1)
        unit = float(it.get("valor_unitario") or it.get("valor") or 0)
        return round(qtd * unit, 2)

    hon = [(it, total_item(it)) for it in inv.get("honorarios", [])]
    des = [(it, total_item(it)) for it in inv.get("despesas", [])]
    ter = [(it, total_item(it)) for it in inv.get("terceiros", [])]

    s_hon = round(sum(v for _, v in hon), 2)
    s_des = round(sum(v for _, v in des), 2)

    trib = []
    for it in inv.get("tributos", []):
        if it.get("valor"):
            v = round(float(it["valor"]), 2)
        else:
            base_nome = str(it.get("base") or "honorarios").lower()
            base = {"honorarios": s_hon, "honorarios_despesas": s_hon + s_des,
                    "total": s_hon + s_des}.get(base_nome, s_hon)
            if it.get("base_calculo"):
                base = float(it["base_calculo"])
            v = round(base * float(it.get("aliquota") or 0) / 100.0, 2)
        trib.append((it, v))
    s_trib = round(sum(v for _, v in trib), 2)

    bruto = round(s_hon + s_des + s_trib, 2)
    desc_pct = float(inv.get("desconto", {}).get("percentual") or 0)
    desconto = round(bruto * desc_pct / 100.0, 2)
    liquido = round(bruto - desconto, 2)

    conds = d["pagamento"].get("condicoes") or []
    for c in conds:
        c["valor"] = round(liquido * float(c.get("percentual") or 0) / 100.0, 2)
    soma_pct = round(sum(float(c.get("percentual") or 0) for c in conds), 2)

    return {
        "honorarios": hon, "despesas": des, "tributos": trib, "terceiros": ter,
        "soma_honorarios": s_hon, "soma_despesas": s_des, "soma_tributos": s_trib,
        "soma_terceiros": round(sum(v for _, v in ter), 2),
        "bruto": bruto, "desconto_pct": desc_pct, "desconto": desconto,
        "liquido": liquido, "extenso": extenso(liquido),
        "soma_percentuais": soma_pct,
        "por_hectare": round(s_hon / d["imovel"]["area_ha"], 2) if d["imovel"].get("area_ha") else 0,
    }


def nome_arquivo(d: dict) -> str:
    """Padrão exigido: ano_mes_dia_versão.md -> 2026_11_20_v01.md"""
    p = d["proposta"]
    dt: date = p["data"]
    return f"{dt:%Y_%m_%d}_v{int(p['versao']):02d}.md"


def dump_json(obj) -> str:
    return json.dumps(obj, ensure_ascii=False, indent=2, default=str)
