# Planilha de despesas de execução (`despesa_AAAA_MM_DD_vNN.xls`)

Artefato de **custo direto interno**: mostra quanto custa executar o serviço, separado em mão de
obra, material e serviços terceirizados. Não é preço de venda — honorários, margem e tributos
ficam na proposta (bloco `investimento`).

## Quando gerar

- Sempre que o usuário pedir composição de custos, planilha de despesas, orçamento analítico ou
  “quanto custa executar”.
- Automaticamente, junto com a proposta, quando o YAML tiver `custos_execucao` com algum item
  (o `gerar_proposta.py` chama `gerar_despesas.py`). Use `--sem-despesas` para não gerar.

## Estrutura da planilha

Aba **Despesas** (paisagem, ajustada a uma página de largura):

| Bloco | Colunas | Fórmula |
|---|---|---|
| A. Mão de obra | Função · Qtd. profissionais · Diárias · Valor da diária · Encargos (%) · Total | `qtd × diárias × diária × (1 + encargos)` |
| B. Material e insumos | Item · Quantidade · Unidade · Valor unitário · Observação · Total | `qtd × valor unitário` |
| C. Serviços terceirizados | Serviço · Fornecedor · Quantidade · Unidade · Valor unitário · Total | `qtd × valor unitário` |
| Resumo | categoria · total · participação % | `soma dos subtotais` |
| Conciliação | despesas da proposta (item 7.2) · material+terceiros · diferença · mão de obra · R$/ha | referências cruzadas |

Aba **Premissas**: encargos sociais, diárias, material, terceirizados, o que não está incluso e
aviso de uso interno.

As fórmulas ficam **vivas**: alterar quantidade, diária ou valor unitário recalcula subtotais,
total, participação e conciliação.

## Bloco YAML

```yaml
custos_execucao:
  encargos_sociais: 0        # % sobre a mão de obra (CLT: informar; autônomo: 0)
  mao_de_obra:
    - funcao: "Topógrafo"
      qtd: 1
      dias: 4
      valor_diaria: 320
  materiais:
    - item: "Marco de concreto"
      qtd: 22
      unidade: "un"
      valor_unitario: 38
      observacao: "vértices materializados"
  servicos_terceirizados:
    - item: "Locação de veículo 4x4"
      fornecedor: "a definir"
      qtd: 4
      unidade: "diária"
      valor_unitario: 280
  observacoes: ""
```

## Regras de coerência

1. **Mão de obra não entra na proposta como despesa**: ela já está embutida nos honorários. A
   conciliação compara apenas material + terceirizados com o item 7.2 da proposta.
2. Diferença negativa grande entre proposta e planilha indica despesa subestimada — avise o
   responsável técnico antes de publicar.
3. Encargos sociais: use 0 para equipe autônoma/MEI e o percentual real (INSS patronal, FGTS,
   férias, 13º) quando houver vínculo CLT.
4. Equipamentos próprios não são depreciados nesta planilha; se o usuário quiser, crie um item em
   “serviços terceirizados” chamado “depreciação/locação de equipamentos” com valor por diária.
5. Diárias de campo já incluem hospedagem/alimentação — evite duplicar em terceirizados.
6. Não envie a planilha ao cliente sem revisão: é documento de custo interno.

## Comando isolado

```bash
python3 scripts/gerar_despesas.py dados.yaml --saida ./saida_proposta
```

Retorna JSON com `soma_mao_de_obra`, `soma_materiais`, `soma_terceiros`, `total`,
`despesas_proposta` e `diferenca` — use esses números na resposta ao usuário, sem recalcular.
