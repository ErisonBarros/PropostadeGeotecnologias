# Propostas_MD — arquivo das propostas

Cópia de arquivo, em Markdown, de cada proposta técnica emitida pela R. Oliveira Consultoria.
O documento de envio ao cliente é o `.docx`/`.pdf`; o `.md` existe para consulta, revisão e
versionamento (permite diff entre versões e busca por texto). A planilha `despesa_*.xls` é o
custo de execução, de uso interno.

## Convenção de nomes

Todos os artefatos de uma mesma emissão compartilham a base `AAAA_MM_DD_vNN` — data da proposta
(não a data em que o arquivo foi salvo) e versão com dois dígitos.

| Artefato | Nome |
|---|---|
| Documento Word | `Proposta_AAAA_MM_DD_vNN.docx` |
| PDF | `Proposta_AAAA_MM_DD_vNN.pdf` |
| Markdown | `Proposta_AAAA_MM_DD_vNN.md` |
| Planilha de despesas | `despesa_AAAA_MM_DD_vNN.xls` |

Exemplos: `Proposta_2026_11_20_v01.docx` e `despesa_2026_11_20_v01.xls` (emissão);
`Proposta_2026_11_20_v02.md` (revisão de valor no mesmo dia);
`Proposta_2026_12_03_v01.docx` (nova proposta em outra data).

## Regras

- Revisão de valor, prazo ou escopo **sempre** gera nova versão — nunca sobrescreva a anterior.
- Duas propostas distintas na mesma data: use versões diferentes (`v01`, `v02`), mantendo
  `proposta.numero` distinto no conteúdo.
- Cada Markdown começa com um bloco `<!-- meta: {...} -->` com número, data, versão, cliente,
  objeto, local, área, honorários, valor global e status; é dele que o `INDICE.md` é montado.
- O índice `INDICE.md` é gerado automaticamente — não edite à mão.
- O histórico de precificação fica em `referencias/precos-observados.md`, na raiz do repositório.

## Origem

Gerado pela skill `proposta-geotecnologias` (espelhada em `skill/proposta-geotecnologias/`):
YAML de dados → DOCX com timbre + PDF + Markdown + planilha de despesas → arquivo no Google Drive
(pasta `Propostas_MD`) e neste repositório.
