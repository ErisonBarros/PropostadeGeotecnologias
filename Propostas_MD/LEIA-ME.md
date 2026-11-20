# Propostas_MD — arquivo das propostas

Cópia de arquivo, em Markdown, de cada proposta técnica emitida pela R. Oliveira Consultoria.
O documento de envio ao cliente é o `.docx`/`.pdf`; o `.md` existe para consulta, revisão e
versionamento (permite diff entre versões e busca por texto).

## Convenção de nomes

```
AAAA_MM_DD_vNN.md
```

| Parte | Significado |
|---|---|
| `AAAA_MM_DD` | data da proposta (não a data em que o arquivo foi salvo) |
| `vNN` | versão com dois dígitos: `v01`, `v02`, `v03`... |

Exemplos: `2026_11_20_v01.md` (emissão), `2026_11_20_v02.md` (revisão de valor no mesmo dia),
`2026_12_03_v01.md` (nova proposta em outra data).

## Regras

- Revisão de valor, prazo ou escopo **sempre** gera nova versão — nunca sobrescreva a anterior.
- Cada arquivo começa com um bloco `<!-- meta: {...} -->` com número, data, versão, cliente,
  objeto, local, área, honorários, valor global e status; é dele que o `INDICE.md` é montado.
- O índice `INDICE.md` é gerado automaticamente — não edite à mão.
- O histórico de precificação fica em `referencias/precos-observados.md`, na raiz do repositório.

## Origem

Gerado pela skill `proposta-geotecnologias` (espelhada em `skill/proposta-geotecnologias/`):
YAML de dados → DOCX com timbre + PDF + Markdown → arquivo no Google Drive (pasta `Propostas_MD`)
e neste repositório.
