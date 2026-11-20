---
name: proposta-geotecnologias
description: >
  Especializado na criação automática de propostas técnicas para serviços de geotecnologias,
  gerando um arquivo .docx profissional pronto para download. Use esta skill SEMPRE que o
  usuário solicitar propostas, orçamentos ou documentos técnicos relacionados a:
  georreferenciamento de imóveis rurais, levantamentos topográficos, aerolevantamentos com
  drones, cadastro técnico imobiliário, mapeamento GNSS/RTK, fotogrametria, geoprocessamento,
  ou qualquer serviço de engenharia cartográfica. Também acione quando o usuário fornecer
  dados de cliente, área em hectares, tipo de serviço e responsável técnico — mesmo em
  texto livre — e pedir para gerar uma proposta, orçamento ou documento técnico.
---

# 🧠 SKILL: Gerador de Proposta Técnica em Geotecnologias

## 🎯 Objetivo
Extrair dados de texto livre fornecido pelo usuário, calcular os valores automaticamente e
gerar um arquivo `.docx` profissional no formato **Proposta Técnica**, pronto para download.

---

## 📥 Extração de Dados (Texto Livre)

O usuário pode fornecer os dados em qualquer formato textual. Extraia e mapeie para:

| Campo | Exemplos de como o usuário pode informar |
|---|---|
| `nome_cliente` | "cliente: João Silva", "para: Fazenda São Pedro" |
| `cpf_cnpj` | CPF ou CNPJ mencionado |
| `tipo_servico` | "georreferenciamento", "levantamento topográfico", "drone", "aerolevantamento" |
| `municipio` | cidade/estado mencionado |
| `area_hectares` | "50 ha", "120 hectares", "área de 30ha" |
| `valor_por_hectare` | "R$ 80/ha", "80 reais por hectare" |
| `dias_trabalho` | "5 dias", "prazo de 3 dias" |
| `custo_equipe_dia` | "equipe custa R$ 500/dia", "diária de R$ 800" |
| `custo_deslocamento` | "deslocamento: R$ 300", "combustível R$ 200" |
| `responsavel_tecnico` | "responsável: Eng. Rodrigo", "RT: José" |

**Se algum campo obrigatório estiver ausente**, pergunte apenas o que falta antes de gerar.

---

## 🧮 Cálculos Obrigatórios

Execute antes de gerar o documento:

```
custo_levantamento = area_hectares × valor_por_hectare
custo_equipe       = dias_trabalho × custo_equipe_dia
custo_total        = custo_levantamento + custo_equipe + custo_deslocamento
```

Formate todos os valores monetários como `R$ 0,00` (padrão brasileiro).

---

## 🧾 Geração Automática da Descrição Técnica

Com base em `tipo_servico` e `municipio`, gere um texto técnico formal (3-5 parágrafos) que inclua:
- Objetivo do levantamento
- Metodologia aplicada (equipamentos, normas — ex: GNSS RTK, INCRA, SIGEF)
- Produtos a serem entregues (ex: planta georreferenciada, shapefile, memorial descritivo)
- Referências normativas pertinentes

---

## 📄 Geração do Arquivo .docx

Use `docx` (npm) para criar o documento. Siga rigorosamente o template abaixo.

### Setup
```bash
npm install -g docx 2>/dev/null || true
```

### Estrutura do Documento

```javascript
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
        VerticalAlign } = require('docx');
const fs = require('fs');

// Dados preenchidos após extração e cálculo
const dados = { /* ... */ };

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      // 1. Título
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "PROPOSTA TÉCNICA", bold: true, size: 36 })]
      }),
      new Paragraph({ children: [new TextRun("")] }),

      // 2. Seção 1 - Identificação
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("1. IDENTIFICAÇÃO")] }),
      // Tabela de identificação (4 linhas x 2 colunas)

      // 3. Seção 2 - Parâmetros do Projeto
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("2. PARÂMETROS DO PROJETO")] }),
      // Tabela de parâmetros (5 linhas x 2 colunas)

      // 4. Seção 3 - Planilha de Cálculo
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("3. PLANILHA DE CÁLCULO")] }),
      // Tabela de cálculo (4 linhas x 3 colunas) com total em negrito

      // 5. Seção 4 - Descrição dos Serviços
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("4. DESCRIÇÃO DOS SERVIÇOS")] }),
      // Parágrafos com texto técnico gerado

      // 6. Seção 5 - Prazo
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("5. PRAZO")] }),
      new Paragraph({ children: [new TextRun(`Prazo estimado: ${dados.dias_trabalho} dias corridos a partir da assinatura.`)] }),

      // 7. Seção 6 - Valor Total
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("6. VALOR TOTAL")] }),
      new Paragraph({ children: [new TextRun({ text: `Valor global da proposta: ${dados.custo_total_fmt}`, bold: true })] }),

      // 8. Seção 7 - Responsabilidade Técnica
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("7. RESPONSABILIDADE TÉCNICA")] }),
      new Paragraph({ children: [new TextRun(`Responsável Técnico: ${dados.responsavel_tecnico}`)] }),

      // 9. Seção 8 - Considerações Finais
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("8. CONSIDERAÇÕES FINAIS")] }),
      new Paragraph({ children: [new TextRun("Esta proposta foi elaborada com base nos dados fornecidos pelo cliente e tem validade de 30 (trinta) dias.")] }),

      // 10. Seção 9 - Assinatura
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("9. ASSINATURA")] }),
      new Paragraph({ children: [new TextRun("__________________________________")] }),
      new Paragraph({ children: [new TextRun(dados.responsavel_tecnico)] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/mnt/user-data/outputs/Proposta_Tecnica.docx", buffer);
  console.log("Documento gerado com sucesso.");
});
```

### Tabelas: Use sempre `WidthType.DXA` e `ShadingType.CLEAR`

**Cabeçalho de tabela (header row):**
```javascript
new TableRow({
  tableHeader: true,
  children: [
    new TableCell({
      borders, margins: cellMargins,
      shading: { fill: "1F4E79", type: ShadingType.CLEAR },
      children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Coluna", bold: true, color: "FFFFFF" })] })]
    })
  ]
})
```

**Linha de total (destaque):**
```javascript
new TableRow({
  children: [
    new TableCell({
      borders, margins: cellMargins,
      shading: { fill: "D6E4F0", type: ShadingType.CLEAR },
      children: [new Paragraph({ children: [new TextRun({ text: "TOTAL GERAL", bold: true })] })]
    }),
    // ...
  ]
})
```

---

## 🔄 Fluxo de Execução

1. Extrair todos os campos do texto livre do usuário
2. Identificar campos faltantes e perguntar apenas o necessário
3. Executar os cálculos
4. Gerar a descrição técnica automaticamente
5. Criar o arquivo `.docx` usando `docx` (npm)
6. Salvar em `/mnt/user-data/outputs/Proposta_Tecnica.docx`
7. Usar `present_files` para disponibilizar o download

---

## 🚫 Restrições

- Não inventar dados que não foram fornecidos
- Não usar unicode bullets — usar `LevelFormat.BULLET`
- Não usar `\n` dentro de TextRun — usar Paragraphs separados
- Não usar `WidthType.PERCENTAGE` em tabelas
- Não omitir nenhuma seção do template
- Formatar valores como `R$ 0,00` (vírgula decimal, ponto milhar)
