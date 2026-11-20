# Layout e identidade — folha modelo R. Oliveira Consultoria

Especificação extraída do arquivo `PROPOSTA_FOLHA MODELO.docx` do repositório
`ErisonBarros/PropostadeGeotecnologias` (folha modelo = papel timbrado; o corpo é livre).
O script `scripts/gerar_proposta.py` já implementa tudo abaixo — consulte este arquivo apenas
para alterar a arte, conferir medidas ou explicar o padrão ao usuário.

## Papel e margens

| Item | Valor |
|---|---|
| Papel | A4 — 21,0 × 29,7 cm |
| Margens | 2,0 cm (laterais), 4,0 cm (topo), 3,35 cm (base) |
| Distância do cabeçalho | 0 cm (arte sangrada na borda superior) |
| Distância do rodapé | 2,35 cm (numeração acima da faixa de contato) |

## Arte do timbre

| Elemento | Arquivo | Dimensão na página | Cor dominante |
|---|---|---|---|
| Faixa superior (logotipo R / OLIVEIRA — Regularização de Imóveis) | `templates/identidade/cabecalho.png` | 21,0 × 3,01 cm, ancorada em (0, 0) | azul-marinho `#050A30` |
| Faixa inferior (contatos) | `templates/identidade/rodape.png` | 21,0 × 2,18 cm, ancorada na base | azul `#03406D` |
| Logotipo quadrado (uso avulso: capas, apresentações) | `templates/identidade/logo.png` | livre | grafite `#333333` |

As duas faixas são inseridas como **imagem ancorada relativa à página** (`wp:anchor`), não como
imagem inline — é o que garante sangria total em Word e LibreOffice. Se substituir a arte,
mantenha a proporção 1298:186 (cabeçalho) e 1300:135 (rodapé) para não distorcer o timbre.

Contatos impressos na faixa inferior (não repetir no corpo sem necessidade):
Praça Pedro Coutinho, 56 — Anexo Aderita Rosa de Oliveira, Loja 02, Centro, Paudalho/PE,
CEP 55.825-000 · (81) 98157-1227 · rodrigocabral@uol.com.br · @r.oliveiraconsultoria.

## Tipografia e cores

| Uso | Fonte | Tamanho | Cor |
|---|---|---|---|
| Corpo | Calibri | 10,5 pt, entrelinha 1,12 | `#2B2B2B` |
| Título da proposta | Calibri bold | 16 pt, caixa alta, centralizado | `#050A30` |
| Subtítulo | Calibri itálico | 10,5 pt, centralizado | `#5A6472` |
| Títulos de seção (`1. IDENTIFICAÇÃO`) | Calibri bold | 11,5 pt, filete inferior | `#050A30` |
| Subtítulos (`7.1 Honorários`) | Calibri bold | 10,5 pt | `#03406D` |
| Cabeçalho de tabela | Calibri bold | 9,5 pt, fundo `#050A30` | branco |
| Corpo de tabela | Calibri | 9,5 pt, zebra `#F2F4F7` | `#2B2B2B` |
| Linha de total | Calibri bold, fundo `#DCE6F1` | 9,5 pt | `#050A30` |
| Bordas de tabela | — | 0,5 pt | `#C3CBD6` |
| Notas de rodapé de tabela | Calibri itálico | 8,5 pt | `#5A6472` |

## Estrutura obrigatória do documento (11 seções)

1. Identificação — cliente, documento, endereço, contato, objeto, imóvel, local, data, validade, versão
2. Apresentação — carta de 1 a 3 parágrafos
3. Objeto — resumo, finalidade e (3.1) dados do imóvel
4. Escopo técnico — 4.1 etapas numeradas, 4.2 entregáveis, 4.3 itens não inclusos
5. Metodologia e normas — parágrafos técnicos + 5.1 referências normativas
6. Cronograma — tabela fase/descrição/prazo + prazo total
7. Investimento — 7.1 honorários, 7.2 despesas, 7.3 tributos, 7.4 terceiros (informativos), 7.5 resumo
8. Condições de pagamento — tabela etapa/%/valor
9. Premissas, responsabilidades e exclusões
10. Validade e disposições gerais
11. Responsabilidade técnica e assinatura — texto da ART + blocos contratada/contratante

Não altere a ordem nem remova seções: o arquivo Markdown de arquivo e o DOCX devem manter
numeração idêntica, pois são comparados em revisões.

## Regras de composição

- Valores monetários sempre em `R$ 0.000,00` (ponto de milhar, vírgula decimal) e, no total,
  acompanhados do valor por extenso entre parênteses.
- Honorários, despesas, tributos e taxas de terceiros **nunca** somados em um único bloco:
  terceiros são informativos e ficam fora do valor global.
- Toda tabela tem cabeçalho repetido na quebra de página (`w:tblHeader`) e linhas indivisíveis
  (`w:cantSplit`).
- Não usar marcadores Unicode "soltos" no DOCX: o script usa parágrafos com recuo pendente,
  garantindo alinhamento em qualquer versão do Word.
- PDF é gerado por LibreOffice (`soffice --convert-to pdf`); se o ambiente não tiver LibreOffice,
  entregue DOCX + MD e informe o usuário.
