---
name: proposta-geotecnologias
description: >
  Gera propostas técnicas e comerciais de geotecnologias no layout R. Oliveira (DOCX com timbre,
  PDF e cópia em Markdown), calcula honorários a partir da tabela de referência 2026, valida os
  dados e arquiva a proposta no Google Drive (pasta Propostas_MD) e no repositório GitHub
  PropostadeGeotecnologias. Use SEMPRE que o usuário pedir proposta, orçamento, cotação ou
  documento comercial para georreferenciamento rural (SIGEF/INCRA), levantamento topográfico,
  aerolevantamento com drone, fotogrametria, cadastro técnico imobiliário, REURB, retificação de
  área, desmembramento, remembramento, regularização de edificação (Habite-se), perícia judicial,
  loteamento ou qualquer serviço de engenharia cartográfica e agrimensura — inclusive quando os
  dados vierem em texto livre, planilha ou áudio transcrito.
---

# Gerador de propostas de geotecnologias

Automatiza o ciclo completo: **coleta → cálculo → redação → DOCX/PDF/MD → validação → arquivo no
Drive e no GitHub**. A fonte única de verdade é um arquivo YAML; DOCX, PDF e Markdown saem dele,
portanto nunca ficam divergentes.

## Fluxo obrigatório

1. **Coletar** os dados que faltam (seção *Coleta*).
2. **Precificar** com `references/honorarios-2026.md` (nunca inventar valores).
3. **Redigir** escopo e metodologia com `references/catalogo-servicos.md`.
4. **Gerar** DOCX + PDF + MD com `scripts/gerar_proposta.py`.
5. **Validar** com `scripts/validar_proposta.py` (erros bloqueiam a entrega).
6. **Confirmar** o valor global com o profissional antes de publicar.
7. **Publicar** no Drive e no GitHub com `scripts/publicar.py`.
8. **Entregar** os links dos arquivos e o resumo do que foi arquivado.

## Coleta de dados

Monte o YAML a partir de `templates/proposta.yaml`. Pergunte **somente** o que faltar, em um único
bloco de perguntas. Obrigatórios para gerar:

| Campo | Observação |
|---|---|
| `cliente.nome` | pessoa física ou jurídica |
| `cliente.documento` | CPF/CNPJ (avise se ausente) |
| `objeto.resumo` | descrição objetiva do serviço |
| `investimento.honorarios` | ao menos um item com quantidade e valor unitário |
| `pagamento.condicoes` | percentuais que somam 100% (padrão 40/30/30) |
| `responsavel_tecnico.nome` e `crea` | para a ART e a assinatura |

Recomendados: `imovel` (denominação, município/UF, matrícula, área em ha, perímetro, vértices),
`escopo.entregaveis`, `escopo.nao_inclusos`, `metodologia.normas`, `cronograma`, `premissas`.
Se o serviço é cobrado por hectare, `imovel.area_ha` passa a ser obrigatório.

Quando o usuário não informar um dado, **deixe vazio e registre a pendência** em `_interno.pendencias`
— não preencha com suposição.

## Precificação

Leia `references/honorarios-2026.md` antes de calcular. Regras:

- Georreferenciamento rural: calcule pelos dois regimes (faixa por hectare para áreas pequenas e
  taxa decrescente por hectare com mínimo de serviço para áreas maiores) e adote o **maior**.
- Separe sempre **honorários**, **despesas operacionais**, **tributos** (ISSQN) e **taxas de
  terceiros**. Terceiros são informativos e ficam **fora** do valor global.
- Desconto acima de 8% exige aprovação explícita; o validador emite aviso.
- A ART está embutida nos honorários — declare isso na seção 11.
- Consulte `referencias/precos-observados.md` do repositório para calibrar com o histórico real.

## Redação técnica

Use `references/catalogo-servicos.md` para escolher o bloco de etapas, entregáveis e normas do
serviço. Escreva em português formal, técnico e verificável: sem promessa de aprovação em órgão
público, cartório ou certificação; sem norma inventada (cite apenas as listadas ou confirme antes).

## Geração

```bash
python3 scripts/gerar_proposta.py dados.yaml --saida ./saida_proposta
```

Produz, em `./saida_proposta`:

- `AAAA_MM_DD_vNN.md` — cópia de arquivo (nome definido por `proposta.data` + `proposta.versao`);
- `Proposta_<numero>_<cliente>_AAAA-MM-DD_vNN.docx` — layout com timbre;
- `Proposta_..._vNN.pdf` — via LibreOffice, quando disponível.

O script cria sozinho um venv com `python-docx` e `PyYAML` na primeira execução; não instale nada
manualmente. Retorna um JSON com valor global, extenso, somas e nome do MD — use esses números na
resposta ao usuário, sem recalcular.

## Validação

```bash
python3 scripts/validar_proposta.py dados.yaml --md saida_proposta/AAAA_MM_DD_vNN.md
```

Corrija todos os itens de `erros` (bloqueiam) e avalie os `avisos`. Só siga para publicação com
`"ok": true` e valor confirmado pelo responsável técnico.

## Publicação

```bash
python3 scripts/publicar.py \
  --md saida_proposta/2026_11_20_v01.md \
  --docx saida_proposta/Proposta_....docx \
  --pdf saida_proposta/Proposta_....pdf \
  --aprendizado "45 ha, acesso difícil; fechado em R$ 150/ha" \
  --status enviada
```

- **Drive:** envia o `.md` (e DOCX/PDF) para a pasta `Propostas_MD` dentro de
  `1Yqne1IwJO6AGC0Ng3_NGwQ3PNZU-2waR` (gerador-propostas-geoespaciais), criando a pasta se não existir.
- **GitHub:** copia o `.md` para `Propostas_MD/`, regenera `Propostas_MD/INDICE.md`, acrescenta a
  linha em `referencias/precos-observados.md` e faz commit/push no branch `erison.barros`.
- Use `--sem-drive` ou `--sem-github` quando o conector não estiver disponível; informe o usuário
  do que não foi arquivado.
- `--sincronizar-skill` espelha a skill instalada em `skill/proposta-geotecnologias` do repositório.

Detalhes, convenção de nomes e solução de problemas: `references/publicacao.md`.

## Entrega ao usuário

Responda com: valor global e por extenso, resumo do escopo, prazo, condição de pagamento, links
absolutos do DOCX/PDF/MD, e o resultado do arquivamento (Drive e commit). Liste as pendências que
dependem do cliente. Nunca afirme que algo foi publicado sem o retorno do `publicar.py`.

## Regras rígidas

- Não invente preços, áreas, matrículas, normas, prazos ou dados de cliente.
- Não garanta aprovação, certificação, Habite-se ou registro — explicite a dependência de terceiros.
- Não misture honorários, despesas, tributos e taxas de terceiros em um único total.
- Não altere a ordem das 11 seções nem remova seções do layout.
- Não publique proposta com valor não confirmado pelo profissional.
- Não edite arquivos dentro de `Propostas/` do repositório (acervo histórico).

## Manutenção e aprendizado

Ao final de cada ciclo, compare o valor praticado com `references/honorarios-2026.md`:

1. desvio sistemático para baixo em 2+ propostas do mesmo serviço → proponha revisão de faixa;
2. proposta aprovada acima da faixa → registre como novo teto;
3. atualize `honorarios-2026.md` e a planilha de origem, publique com a mensagem
   `honorarios: revisão de faixa — <serviço>` e mantenha `--sincronizar-skill` ativo para versionar
   a skill no repositório.

Arquivos de apoio: `references/layout-identidade.md` (medidas, cores, tipografia),
`references/catalogo-servicos.md` (escopo por serviço), `references/publicacao.md` (Drive/GitHub),
`templates/proposta.yaml` (esquema), `templates/exemplo-georreferenciamento.yaml` (exemplo completo),
`templates/identidade/` (arte do timbre).
