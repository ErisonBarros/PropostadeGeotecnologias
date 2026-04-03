---

# 🧠 📄 SKILL EM MARKDOWN (PADRÃO CLAUDE)

Copie exatamente como está 👇

````markdown
# 🧠 SKILL: Gerador de Proposta Técnica

## 🎯 Objetivo
Gerar automaticamente um documento técnico profissional no formato **Proposta_Folha_modelo**, com base em dados estruturados fornecidos pelo usuário.

---

## ⚙️ Regras Obrigatórias

- Seguir rigorosamente o layout definido
- Não alterar títulos ou ordem das seções
- Não remover nenhuma seção
- Preencher todas as tabelas obrigatoriamente
- Realizar todos os cálculos antes de gerar o documento
- Não inventar dados
- Usar exclusivamente os dados fornecidos
- Formatar valores monetários como `R$ 0.00`
- Retornar apenas o documento final (sem explicações)

---

## 📥 Entrada Esperada (JSON)

O usuário fornecerá os dados no seguinte formato:

```json
{
  "nome_cliente": "",
  "cpf_cnpj": "",
  "tipo_servico": "",
  "municipio": "",
  "area_hectares": 0,
  "valor_por_hectare": 0,
  "dias_trabalho": 0,
  "custo_equipe_dia": 0,
  "custo_deslocamento": 0,
  "responsavel_tecnico": ""
}
````

---

## 🧮 Regras de Cálculo

Calcular obrigatoriamente:

* custo_levantamento = area_hectares × valor_por_hectare
* custo_equipe = dias_trabalho × custo_equipe_dia
* custo_total = custo_levantamento + custo_equipe + custo_deslocamento

---

## 🧾 Geração da Descrição Técnica

Gerar automaticamente um texto técnico profissional com base em:

* tipo_servico
* municipio

O texto deve ser:

* Formal
* Objetivo
* Técnico
* Coerente com topografia/cartografia

---

## 📄 Template do Documento (OBRIGATÓRIO)

```markdown
# PROPOSTA TÉCNICA

---

## 1. IDENTIFICAÇÃO

Cliente: {nome_cliente}  
CPF/CNPJ: {cpf_cnpj}  
Município: {municipio}  
Serviço: {tipo_servico}  

Responsável Técnico: {responsavel_tecnico}

---

## 2. PARÂMETROS DO PROJETO

| Item                    | Valor |
|-------------------------|------|
| Área (ha)              | {area_hectares} |
| Valor por hectare (R$) | {valor_por_hectare} |
| Dias de trabalho       | {dias_trabalho} |
| Custo equipe/dia (R$)  | {custo_equipe_dia} |
| Deslocamento (R$)      | {custo_deslocamento} |

---

## 3. PLANILHA DE CÁLCULO

| Descrição              | Fórmula                         | Valor (R$) |
|------------------------|---------------------------------|------------|
| Levantamento          | área × valor/ha                | {custo_levantamento} |
| Equipe                | dias × custo/dia               | {custo_equipe} |
| Deslocamento          | valor fixo                     | {custo_deslocamento} |
| **TOTAL GERAL**       | soma                           | **{custo_total}** |

---

## 4. DESCRIÇÃO DOS SERVIÇOS

{descricao_servico}

---

## 5. PRAZO

Prazo estimado: {dias_trabalho} dias.

---

## 6. VALOR TOTAL

Valor global da proposta: **R$ {custo_total}**

---

## 7. RESPONSABILIDADE TÉCNICA

Responsável Técnico: {responsavel_tecnico}

---

## 8. CONSIDERAÇÕES FINAIS

Esta proposta foi elaborada com base nos dados fornecidos pelo cliente.

---

## 9. ASSINATURA

__________________________________  
{responsavel_tecnico}
```

---

## 🔄 Fluxo de Execução

1. Validar se todos os campos foram fornecidos
2. Executar todos os cálculos
3. Gerar descrição técnica
4. Preencher o template
5. Retornar apenas o documento final

---

## 🚫 Restrições

* Não adicionar comentários
* Não explicar o processo
* Não modificar o template
* Não omitir valores
* Não usar placeholders vazios

---

## ✅ Saída Esperada

A saída deve ser **apenas o documento final em Markdown**, totalmente preenchido.

---

