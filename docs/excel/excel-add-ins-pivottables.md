---
title: Trabalhar com tabelas dinâmicas usando a API JavaScript do Excel
description: Use a API do JavaScript Excel para criar tabelas dinâmicas e interagir com seus componentes.
ms.date: 09/21/2018
ms.openlocfilehash: 00dd982d4ba4de0db34277cd546b572d4394e258
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459277"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Trabalhar com tabelas dinâmicas usando a API JavaScript do Excel

As tabelas dinâmicas simplificam os conjuntos de dados maiores. Permitem a manipulação rápida de dados agrupados. A API JavaScript do Excel possibilita que os suplementos criem tabelas dinâmicas e interajam com seus componentes. 

Se não está familiarizado com a funcionalidade das tabelas dinâmicas, considere explorá-las como usuário final. Consulte [Criar uma tabela dinâmica para analisar dados de planilhas](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) para obter uma boa orientação sobre essas ferramentas. 

Este artigo fornece exemplos de código para cenários comuns. Para enriquecer a compreensão da API de tabela dinâmica, consulte [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) e [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).

> [!IMPORTANT]
> As tabelas dinâmicas criadas com OLAP não são suportadas no momento.

## <a name="hierarchies"></a>Hierarquias

As tabelas dinâmicas são organizadas com base em quatro categorias de hierarquia: linha, coluna, dados e filtro. Os dados a seguir, que descrevem as vendas de frutas de várias fazendas, serão utilizados ao longo deste artigo.

![Conjunto das vendas de fruta de diferentes tipos provenientes de várias fazendas.](../images/excel-pivots-raw-data.png)

Esses dados têm cinco hierarquias: **Fazendas**, **Tipo**, **Classificação**, **Caixas vendidas na fazenda**, e **Caixas vendidas por atacado**. Cada hierarquia só pode existir em uma das quatro categorias. Se **Tipo** for adicionado as hierarquias de coluna e depois adicionado as hierarquias de linha, ele permanecerá apenas no último.

As hierarquias de linha e coluna definem como os dados serão agrupados. Por exemplo, uma hierarquia de linha de **Fazendas** agrupará todos os conjuntos de dados da mesma fazenda. A escolha entre hierarquia de linha e coluna define a orientação da tabela dinâmica.

As hierarquias de dados são os valores a serem agregados com base nas hierarquias de linhas e colunas. Uma tabela dinâmica com a hierarquia de linhas **Fazendas** e a hierarquia de dados  **Caixas vendidas por atacado** mostra a soma total (por padrão) de todas as frutas diferentes para cada fazenda.

As hierarquias de filtro incluem ou excluem dados do pivô com base nos valores desse tipo filtrado. Uma hierarquia de filtro de **Classificação** com o tipo **Orgânico** selecionado mostra apenas os dados para fruta orgânica.

Aqui estão os dados da fazenda novamente, junto com uma tabela dinâmica. A tabela dinâmica está usando **Fazenda** e **Tipo** como as hierarquias de linha, **Caixas vendidas na fazenda** e **Caixas vendidas por atacado** como as hierarquias de dados (com a função de agregação de soma padrão) e **Classificação** como uma hierarquia de filtro (com **Orgânico** selecionado). 

![Uma seleção de dados de vendas de frutas ao lado de uma tabela dinâmica com hierarquias de linhas, dados e filtros.](../images/excel-pivot-table-and-data.png)

Esta tabela dinâmica pode ser gerada por meio da API do JavaScript ou da interface gráfica do Excel. Ambas as opções permitem mais manipulação através de suplementos.

## <a name="create-a-pivottable"></a>Criar uma tabela dinâmica

Tabelas dinâmicas precisam de um nome, origem e destino. A origem pode ser um endereço de intervalo ou um nome de tabela  (transmitido como um tipo `Range`, `string` ou `Table` ). O destino é um endereço de intervalo (fornecido como `Range` ou `string`). Os exemplos a seguir mostram várias técnicas de criação de tabelas dinâmicas.

### <a name="create-a-pivottable-with-range-addresses"></a>Criar uma tabela dinâmica com endereços de intervalo

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>Criar uma tabela dinâmica com objetos de intervalo

```typescript
await Excel.run(async (context) => {    
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21
    const rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    const rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add(
        "Farm Sales", rangeToAnalyze, rangeToPlacePivot);
    
    await context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a>Criar uma tabela dinâmica no nível da pasta de trabalho

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>Usar uma tabela dinâmica existente

As tabelas dinâmicas criadas manualmente também são acessíveis através da coleção de tabela dinâmica da pasta de trabalho ou de planilhas individuais. 

O código a seguir obtém a primeira tabela dinâmica na pasta de trabalho. Em seguida, fornece um nome para a tabela para facilitar a referência futura.

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>Adicionar linhas e colunas à tabela dinâmica

As linhas e colunas articulam os dados em torno dos valores desses campos.

Adicionar a coluna **Fazenda** articula todas as vendas ao redor de cada fazenda. Adicionar as linhas **Tipo** e **Classificação** divide ainda mais os dados com base no tipo de fruta vendida e se a mesma era orgânica ou não.

![Uma tabela dinâmica com a coluna Fazenda e as linhas Tipo e Classificação.](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

Você também pode ter uma tabela dinâmica apenas com linhas ou colunas.

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>Adicionar hierarquias de dados à tabela dinâmica

As hierarquias de dados preenchem a tabela dinâmica com informações para combinar com base nas linhas e colunas. Adicionar as hierarquias de dados de **Caixas vendidas na fazenda** e **Caixas vendidas por atacado** fornece a soma desses números para cada linha e coluna. 

No exemplo, **Fazenda** e **Tipo** são linhas com os dados das vendas de caixas. 

![Uma tabela dinâmica que mostra as vendas totais das diferentes frutas com base na fazenda de onde elas vieram.](../images/excel-pivots-data-hierarchy.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the hierarchies
    // that will have their data aggregated (summed in this case)
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    await context.sync();
});
```

## <a name="change-aggregation-function"></a>Alterar a função de agregação

As hierarquias de dados têm seus valores agregados. Para conjuntos de dados de números, por padrão, isso corresponde a uma soma. Esse comportamento é definido pela propriedade `summarizeBy` com base no tipo `AggregrationFunction` . 

Os tipos de função agregada suportados atualmente são `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP` e `Automatic` (padrão).

O exemplo de código a seguir altera a agregação para as médias dos dados.

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.dataHierarchies.load("no-properties-needed");
    await context.sync();

    // changing the aggregation from the default sum to an average of all the values in the hierarchy
    pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
    pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
    await context.sync();
});
```

## <a name="change-calculations-with-a-showasrule"></a>Altere os cálculos com ShowAsRule

As tabelas dinâmicas, por padrão, agregam os dados de suas hierarquias de linha e coluna de forma independente. Uma `ShowAsRule` altera a hierarquia dos dados para valores de saída com base em outros itens na tabela dinâmica.

O objeto `ShowAsRule` tem três propriedades:
-   `calculation`: O tipo de cálculo relativo para aplicar à hierarquia de dados (o padrão é `none`).
-   `baseField`: O campo dentro da hierarquia que contém os dados de base antes que o cálculo seja aplicado. O `PivotField` geralmente tem o mesmo nome que sua hierarquia pai.
-   `baseItem`: O item individual comparado com os valores dos campos de base de acordo com o tipo de cálculo. Nem todos os cálculos exigem esse campo.

O exemplo a seguir define o cálculo na hierarquia de dados **Soma das caixas vendidas na Fazenda** para uma porcentagem do total da coluna. Ainda queremos que a granularidade se estenda ao nível do tipo de fruta, então usaremos a hierarquia de linha **Tipo** e o campo subjacente. O exemplo também tem **Fazenda** como a primeira hierarquia de linha, de modo que a entrada total da fazenda exibe também a porcentagem que cada fazenda é responsável por produzir.

![Uma tabela dinâmica que mostra as porcentagens de venda de frutas em relação ao total geral, tanto por fazenda quanto por tipo de fruta dentro de cada fazenda.](../images/excel-pivots-showas-percentage.png)

``` TypeScript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    const farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    await context.sync();

    // show the crates of each fruit type sold at the farm as a percentage of the column's total
    let farmShowAs = farmDataHierarchy.showAs;
    farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;
    farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");
    farmDataHierarchy.showAs = farmShowAs; 
    farmDataHierarchy.name = "Percentage of Total Farm Sales";

    await context.sync();
});
```

O exemplo anterior definiu o cálculo para a coluna, relativo a uma hierarquia de linha individual. Quando o cálculo está relacionado a um item individual, use a propriedade `baseItem` . 

O exemplo a seguir mostra o cálculo `differenceFrom` . Exibe a diferença das entradas da hierarquia de dados de vendas de caixas na fazenda em relação  àquelas das "Fazendas A". O `baseField` é **Fazenda**, portanto, vemos as diferenças entre as outras fazendas, bem como as divisões para cada tipo de fruta (**Tipo** também é uma hierarquia de linha neste exemplo).

![Uma Tabela Dinâmica mostrando as diferenças de vendas de frutas entre “Fazendas A” e as outras. Isso mostra a diferença no total de vendas de frutas das fazendas e as vendas de tipos de frutas. Se “Fazendas A” não vendeu um tipo específico de fruta,  é exibida a mensagem “#N/A”.](../images/excel-pivots-showas-differencefrom.png)

``` TypeScript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    const farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    await context.sync();

    // show the difference between crate sales of the "A Farms" and the other farms
    // this difference is both aggregated and shown for individual fruit types (where applicable)
    let farmShowAs = farmDataHierarchy.showAs;
    farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;
    farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");
    farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");
    farmDataHierarchy.showAs = farmShowAs;
    farmDataHierarchy.name = "Difference from A Farms";
    await context.sync();
});
```

## <a name="pivottable-layouts"></a>Layouts de tabela dinâmica

Um layout de tabela dinâmica define o posicionamento de hierarquias e seus dados. Você acessa o layout para determinar os intervalos em que os dados são armazenados. 

O diagrama a seguir mostra as chamadas de funções de layout que correspondem a cada intervalo da tabela dinâmica.

![Um diagrama que mostra quais seções de uma tabela dinâmica são retornadas pelas funções de intervalo do layout.](../images/excel-pivots-layout-breakdown.png)

O código a seguir demonstra como obter a última linha dos dados de tabela dinâmica percorrendo o layout. Esses valores são então somados para obter um total geral.

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // get the totals for each data hierarchy from the layout
    const range = pivotTable.layout.getDataBodyRange();
    const grandTotalRange = range.getLastRow();
    grandTotalRange.load("address");
    await context.sync();

    // sum the totals from the PivotTable data hierarchies and place them in a new range
    const masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("B27:C27");
    masterTotalRange.formulas = [["All Crates", "=SUM(" + grandTotalRange.address + ")"]];
    await context.sync();
});
```

As tabelas dinâmicas tês três estilos de layout: Compacto, Estrutura do Código e Tabular. Nos exemplos anteriores foi usado o estilo compacto. 

Os exemplos a seguir usam os estilos de estrutura de código e tabular, respectivamente. O exemplo de código mostra como alternar entre os diferentes layouts.

### <a name="outline-layout"></a>Layout de estrutura do código

![Uma tabela dinâmica usando o layout de estrutura de tópicos.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a>Layout tabular

![Uma tabela dinâmica usando o layout tabular.](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a>Alterar os nomes de hierarquia

Os campos de hierarquia são editáveis. O código a seguir demonstra como alterar os nomes exibidos de duas hierarquias de dados.

```typescript
await Excel.run(async (context) => {
    const dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    await context.sync();
    
    // changing the displayed names of these entries
    dataHierarchies.items[0].name = "Farm Sales";
    dataHierarchies.items[1].name = "Wholesale";
    await context.sync();
});
```

## <a name="delete-a-pivottable"></a>Excluir uma tabela dinâmica

As tabelas dinâmicas são excluídas pelo uso de seu nome.

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a>Confira também

- [Conceitos básicos de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Referência da API JavaScript do Excel](https://docs.microsoft.com/javascript/api/excel)
