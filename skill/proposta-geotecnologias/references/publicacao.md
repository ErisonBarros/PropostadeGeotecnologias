# Arquivo, versionamento e sincronização

## Convenção de nomes

`AAAA_MM_DD_vNN.md` — data da proposta (não a data de hoje) e versão com dois dígitos.

| Situação | Nome |
|---|---|
| Primeira emissão de 20/11/2026 | `2026_11_20_v01.md` |
| Revisão de valor | `2026_11_20_v02.md` |
| Nova proposta no mesmo dia | `2026_11_20_v02.md` com `proposta.numero` diferente |

A versão vem de `proposta.versao` no YAML. Ao revisar, incremente a versão e mantenha o mesmo
número de proposta — o histórico fica rastreável no Drive e no Git.

## Destinos

| Destino | Caminho | Conteúdo |
|---|---|---|
| Pasta local de trabalho | `./saida_proposta/` | `.docx`, `.pdf`, `.md` |
| Google Drive | pasta `Propostas_MD` dentro de `1Yqne1IwJO6AGC0Ng3_NGwQ3PNZU-2waR` (gerador-propostas-geoespaciais) | `.md` (sempre), `.docx` e `.pdf` (opcional) |
| GitHub | `PropostadeGeotecnologias` (branch `erison.barros`) | `Propostas_MD/AAAA_MM_DD_vNN.md`, `Propostas_MD/INDICE.md`, `referencias/precos-observados.md`, `skill/proposta-geotecnologias/` |

O ID da subpasta é memorizado em `~/.cache/proposta-geo/config.json` após a primeira execução.
Se a pasta for apagada no Drive, remova essa chave do cache para recriar.

## Comando único

```bash
python3 scripts/publicar.py \
  --md   saida_proposta/2026_11_20_v01.md \
  --docx saida_proposta/Proposta_2026-014_cliente_2026-11-20_v01.docx \
  --pdf  saida_proposta/Proposta_2026-014_cliente_2026-11-20_v01.pdf \
  --aprendizado "Área de 45 ha, acesso difícil; fechado em R$ 150/ha" \
  --status enviada
```

Opções úteis: `--sem-drive`, `--sem-github`, `--sem-docx-pdf`, `--status aprovada|perdida`,
`--sincronizar-skill` (copia a skill instalada para `skill/proposta-geotecnologias` no repositório),
`--mensagem` (commit personalizado).

## Estrutura do repositório após a primeira publicação

```
PropostadeGeotecnologias/
├── Propostas_MD/
│   ├── AAAA_MM_DD_vNN.md      ← cópia de arquivo de cada proposta
│   └── INDICE.md              ← tabela gerada automaticamente
├── referencias/
│   └── precos-observados.md   ← memória de precificação (append a cada publicação)
├── skill/proposta-geotecnologias/  ← espelho versionado desta skill
├── Propostas/                 ← acervo histórico original (não alterar)
├── PROPOSTA_FOLHA MODELO.docx ← timbre de origem
└── Honorários_R_Oliveira_2026.xlsx
```

## Ciclo de aprendizado

1. A cada publicação, `publicar.py` acrescenta uma linha em `referencias/precos-observados.md`
   com data, serviço, local, área, honorários, valor global, R$/ha e status.
2. Antes de precificar uma nova proposta, compare a faixa usada com esse histórico
   (ver `honorarios-2026.md`, seção 5).
3. Se o histórico indicar desvio sistemático, atualize `honorarios-2026.md` e a planilha de
   origem, e publique com a mensagem `honorarios: revisão de faixa — <serviço>`.
4. Ao encerrar o ciclo comercial, atualize o status da proposta (`--status aprovada|perdida`),
   reexecutando `publicar.py` com o mesmo arquivo MD — o índice e o histórico são reescritos.

## Verificações e solução de problemas

| Sintoma | Causa provável | Ação |
|---|---|---|
| `gws` não encontrado | Google Workspace CLI ausente | Entregue DOCX/PDF/MD localmente e informe o usuário; ou configure o conector Google |
| `401/403` no Drive | token expirado | Reautenticar o conector Google e repetir `publicar.py` |
| `push falhou` / `403` | credencial do GitHub ausente | `gh auth status`; se necessário, peça ao usuário para reautenticar |
| `non-fast-forward` | repositório atualizado remotamente | o script tenta `pull --rebase --autostash`; se persistir, `rm -rf ~/.cache/proposta-geo/repo` e rodar de novo |
| PDF não gerado | LibreOffice ausente | entregar DOCX + MD e avisar |
| Pasta `Propostas_MD` não encontrada no Drive | apagada ou renomeada | remover a chave `drive_*` de `~/.cache/proposta-geo/config.json` |

Nunca publique uma proposta sem antes validar (`scripts/validar_proposta.py`) e sem confirmar o
valor global com o responsável técnico.
