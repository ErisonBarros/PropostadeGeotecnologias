#!/usr/bin/env python3
"""Arquiva a proposta em Markdown no Google Drive e sincroniza o repositório GitHub.

Uso típico (depois de gerar_proposta.py):
    python3 publicar.py --md saida_proposta/2026_11_20_v01.md \
                        --docx saida_proposta/Proposta_....docx \
                        --pdf  saida_proposta/Proposta_....pdf

O que faz:
  1. garante a pasta "Propostas_MD" na pasta-mãe do Drive (padrão: pasta do projeto);
  2. envia/substitui o .md (e opcionalmente .docx/.pdf) nessa pasta;
  3. clona/atualiza o repositório GitHub, copia o .md para Propostas_MD/,
     regenera Propostas_MD/INDICE.md e registra o aprendizado de precificação
     em referencias/precos-observados.md;
  4. faz commit e push.

Variáveis de ambiente respeitadas: GH_TOKEN / credenciais do gh CLI, GOOGLE_WORKSPACE_CLI_TOKEN.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _comum as C  # noqa: E402

CACHE = Path.home() / ".cache" / "proposta-geo"
CACHE.mkdir(parents=True, exist_ok=True)
CONFIG = CACHE / "config.json"

DRIVE_PAI_PADRAO = "1Yqne1IwJO6AGC0Ng3_NGwQ3PNZU-2waR"  # pasta gerador-propostas-geoespaciais
SUBFOLDER = "Propostas_MD"
REPO_PADRAO = "https://github.com/ErisonBarros/PropostadeGeotecnologias.git"
BRANCH_PADRAO = "erison.barros"


# ------------------------------------------------------------------ utilidades
def carregar_config() -> dict:
    if CONFIG.exists():
        try:
            return json.loads(CONFIG.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def salvar_config(cfg: dict) -> None:
    CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def rodar(cmd: list[str], cwd: Path | None = None, check=True) -> subprocess.CompletedProcess:
    res = subprocess.run(cmd, cwd=str(cwd) if cwd else None, capture_output=True, text=True)
    if check and res.returncode != 0:
        raise RuntimeError(f"falhou: {' '.join(cmd[:3])}...\n{res.stderr[:800]}")
    return res


def gws(args: list[str]) -> dict:
    exe = shutil.which("gws")
    if not exe:
        raise RuntimeError("gws (Google Workspace CLI) não encontrado no ambiente.")
    res = rodar([exe, "drive"] + args)
    saida = (res.stdout or "").strip()
    if not saida:
        return {}
    try:
        return json.loads(saida)
    except json.JSONDecodeError:
        # pode vir NDJSON
        linhas = [json.loads(l) for l in saida.splitlines() if l.strip().startswith("{")]
        return linhas[-1] if linhas else {}


def metadados_md(caminho: Path) -> dict:
    txt = caminho.read_text(encoding="utf-8", errors="ignore")
    m = re.search(r"<!--\s*meta:\s*(\{.*?\})\s*-->", txt, re.S)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass
    return {}


# ------------------------------------------------------------------ Drive
def garantir_subpasta(pai: str, cfg: dict) -> str:
    chave = f"drive_{pai}_{SUBFOLDER}"
    if cfg.get(chave):
        return cfg[chave]
    res = gws(["files", "list", "--params", json.dumps({
        "q": f"name='{SUBFOLDER}' and '{pai}' in parents and "
             f"mimeType='application/vnd.google-apps.folder' and trashed=false",
        "fields": "files(id,name)",
        "supportsAllDrives": "true",
        "includeItemsFromAllDrives": "true",
    })])
    arquivos = res.get("files") or []
    if arquivos:
        fid = arquivos[0]["id"]
    else:
        criado = gws(["files", "create", "--json", json.dumps({
            "name": SUBFOLDER,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [pai],
        })])
        fid = criado.get("id")
        if not fid:
            raise RuntimeError("não foi possível criar a pasta Propostas_MD no Drive")
    cfg[chave] = fid
    salvar_config(cfg)
    return fid


def enviar_drive(arquivo: Path, pasta_id: str) -> dict:
    """Cria ou substitui o arquivo na pasta do Drive (idempotente por nome)."""
    existentes = gws(["files", "list", "--params", json.dumps({
        "q": f"name='{arquivo.name}' and '{pasta_id}' in parents and trashed=false",
        "fields": "files(id,name,size,modifiedTime)",
        "supportsAllDrives": "true",
        "includeItemsFromAllDrives": "true",
    })])
    achados = existentes.get("files") or []
    mime = {".md": "text/markdown", ".docx": "application/vnd.openxmlformats-officedocument"
                                            ".wordprocessingml.document",
            ".pdf": "application/pdf"}.get(arquivo.suffix.lower(), "application/octet-stream")
    if achados:
        fid = achados[0]["id"]
        gws(["files", "update", "--params", json.dumps({"fileId": fid,
                                                        "supportsAllDrives": "true"}),
             "--upload", str(arquivo), "--upload-content-type", mime])
        return {"acao": "atualizado", "id": fid, "nome": arquivo.name}
    criado = gws(["files", "create", "--params", json.dumps({"supportsAllDrives": "true"}),
                  "--upload", str(arquivo), "--upload-content-type", mime,
                  "--json", json.dumps({"name": arquivo.name, "parents": [pasta_id]})])
    return {"acao": "criado", "id": criado.get("id", ""), "nome": arquivo.name}


# ------------------------------------------------------------------ GitHub
def preparar_repo(url: str, branch: str) -> Path:
    destino = CACHE / "repo"
    if not (destino / ".git").exists():
        if destino.exists():
            shutil.rmtree(destino)
        rodar(["git", "clone", "--branch", branch, url, str(destino)])
    else:
        rodar(["git", "remote", "set-url", "origin", url], cwd=destino, check=False)
        rodar(["git", "fetch", "origin", branch], cwd=destino, check=False)
        rodar(["git", "checkout", branch], cwd=destino, check=False)
        rodar(["git", "pull", "--rebase", "--autostash", "origin", branch], cwd=destino,
              check=False)
    return destino


def regerar_indice(repo: Path) -> Path:
    pasta = repo / SUBFOLDER
    pasta.mkdir(parents=True, exist_ok=True)
    linhas = ["# Índice de propostas (Markdown)", "",
              "Arquivo gerado automaticamente pela skill `proposta-geotecnologias`. "
              "Cada proposta é arquivada como `AAAA_MM_DD_vNN.md`.", "",
              "| Arquivo | Data | Versão | Cliente | Serviço / Objeto | Local | Área (ha) | Valor (R$) | Status |",
              "|---|---|---|---|---|---|---|---|---|"]
    itens = []
    for md in sorted(pasta.glob("*.md")):
        if md.name.lower() == "indice.md":
            continue
        m = metadados_md(md)
        itens.append((m.get("data", md.stem), md, m))
    for _, md, m in sorted(itens, key=lambda x: x[0], reverse=True):
        valor = m.get("valor")
        valor_txt = C.brl(valor) if isinstance(valor, (int, float)) else ""
        data_txt = m.get("data", "")
        if data_txt:
            try:
                data_txt = datetime.strptime(data_txt, "%Y-%m-%d").strftime("%d/%m/%Y")
            except ValueError:
                pass
        objeto = str(m.get("objeto") or "")[:90].replace("|", "\\|")
        linha = (f"| [{md.name}]({md.name}) | {data_txt} | v{int(m.get('versao', 1)):02d} | "
                 f"{str(m.get('cliente') or '').replace('|', '/')} | {objeto} | "
                 f"{m.get('municipio') or ''} | "
                 f"{C.num(m['area_ha']) if m.get('area_ha') else ''} | {valor_txt} | "
                 f"{m.get('status') or ''} |")
        linhas.append(linha)
    if len(linhas) == 6:
        linhas.append("| — | — | — | — | Nenhuma proposta arquivada ainda | — | — | — | — |")
    destino = pasta / "INDICE.md"
    destino.write_text("\n".join(linhas) + "\n", encoding="utf-8")
    return destino


def registrar_aprendizado(repo: Path, meta: dict, observacao: str) -> Path:
    pasta = repo / "referencias"
    pasta.mkdir(parents=True, exist_ok=True)
    arq = pasta / "precos-observados.md"
    if not arq.exists():
        arq.write_text(
            "# Preços observados (memória de precificação)\n\n"
            "Registro histórico das propostas emitidas, usado para calibrar as faixas de "
            "honorários da skill. Valores em reais, sem tributos de terceiros.\n\n"
            "| Data | Proposta | Serviço / Objeto | Local | Área (ha) | Honorários (R$) | "
            "Valor global (R$) | R$/ha (honorários) | Status | Observação |\n"
            "|---|---|---|---|---|---|---|---|---|---|\n", encoding="utf-8")
    area = meta.get("area_ha") or 0
    valor = meta.get("valor") or 0
    hon = meta.get("honorarios")
    r_ha = (hon / area) if (hon and area) else (valor / area if area else 0)
    linha = (f"| {meta.get('data','')} | {meta.get('numero','')} | "
             f"{str(meta.get('objeto') or '').replace('|', '/')[:120]} | "
             f"{meta.get('municipio','')} | {C.num(area) if area else ''} | "
             f"{C.brl(hon) if hon else ''} | {C.brl(valor)} | "
             f"{C.num(r_ha) if r_ha else ''} | {meta.get('status','')} | "
             f"{observacao.replace('|', '/')} |")
    with arq.open("a", encoding="utf-8") as fh:
        fh.write(linha + "\n")
    return arq


def commit_push(repo: Path, branch: str, mensagem: str) -> dict:
    rodar(["git", "add", "-A"], cwd=repo)
    status = rodar(["git", "status", "--porcelain"], cwd=repo).stdout.strip()
    if not status:
        return {"commit": None, "push": "nada a enviar"}
    ident = ["-c", "user.name=ErisonBarros",
             "-c", "user.email=ErisonBarros@users.noreply.github.com"]
    rodar(["git"] + ident + ["commit", "-m", mensagem], cwd=repo)
    sha = rodar(["git", "rev-parse", "--short", "HEAD"], cwd=repo).stdout.strip()
    push = rodar(["git", "push", "origin", branch], cwd=repo, check=False)
    ok = push.returncode == 0
    return {"commit": sha, "push": "ok" if ok else f"falhou: {push.stderr[:300]}",
            "branch": branch}


# ------------------------------------------------------------------ CLI
def main() -> None:
    ap = argparse.ArgumentParser(description="Publica a proposta no Drive e no GitHub.")
    ap.add_argument("--md", required=True, help="arquivo Markdown da proposta")
    ap.add_argument("--docx", help="DOCX para enviar ao Drive (opcional)")
    ap.add_argument("--pdf", help="PDF para enviar ao Drive (opcional)")
    ap.add_argument("--drive-pai", default=DRIVE_PAI_PADRAO)
    ap.add_argument("--repo", default=REPO_PADRAO)
    ap.add_argument("--branch", default=BRANCH_PADRAO)
    ap.add_argument("--sem-drive", action="store_true")
    ap.add_argument("--sem-github", action="store_true")
    ap.add_argument("--sem-docx-pdf", action="store_true",
                    help="não enviar DOCX/PDF ao Drive (só o Markdown)")
    ap.add_argument("--aprendizado", default="",
                    help="observação de precificação para referencias/precos-observados.md")
    ap.add_argument("--status", default="", help="status comercial (enviada, aprovada, perdida)")
    ap.add_argument("--mensagem", default="", help="mensagem de commit personalizada")
    ap.add_argument("--sincronizar-skill", action="store_true",
                    help="copia a skill instalada para skill/proposta-geotecnologias no repositório")
    args = ap.parse_args()

    md = Path(args.md).expanduser().resolve()
    if not md.exists():
        raise SystemExit(f"Markdown não encontrado: {md}")
    meta = metadados_md(md)
    if args.status:
        meta["status"] = args.status
    cfg = carregar_config()
    resultado: dict = {"md": str(md), "nome_md": md.name, "meta": meta}

    if not args.sem_drive:
        try:
            pasta = garantir_subpasta(args.drive_pai, cfg)
            enviados = [enviar_drive(md, pasta)]
            if not args.sem_docx_pdf:
                for extra in (args.docx, args.pdf):
                    if extra and Path(extra).exists():
                        enviados.append(enviar_drive(Path(extra).resolve(), pasta))
            resultado["drive"] = {"pasta_id": pasta, "arquivos": enviados,
                                  "link": f"https://drive.google.com/drive/folders/{pasta}"}
        except Exception as exc:  # noqa: BLE001
            resultado["drive"] = {"erro": str(exc)[:500]}

    if not args.sem_github:
        try:
            repo = preparar_repo(args.repo, args.branch)
            destino = repo / SUBFOLDER / md.name
            destino.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(md, destino)
            indice = regerar_indice(repo)
            aprend = registrar_aprendizado(repo, meta, args.aprendizado)
            skill_info = None
            if args.sincronizar_skill:
                alvo = repo / "skill" / "proposta-geotecnologias"
                if alvo.exists():
                    shutil.rmtree(alvo)
                shutil.copytree(C.SKILL_DIR, alvo,
                                ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".git"))
                skill_info = str(alvo.relative_to(repo))
            msg = args.mensagem or (
                f"proposta: {md.name} — {meta.get('cliente','cliente')} "
                f"({C.brl(meta.get('valor', 0))})")
            info = commit_push(repo, args.branch, msg)
            resultado["github"] = {
                "repo": args.repo, "branch": args.branch,
                "arquivo": str(destino.relative_to(repo)),
                "indice": str(indice.relative_to(repo)),
                "aprendizado": str(aprend.relative_to(repo)),
                "skill": skill_info,
                **info,
            }
        except Exception as exc:  # noqa: BLE001
            resultado["github"] = {"erro": str(exc)[:800]}

    print(json.dumps(resultado, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
