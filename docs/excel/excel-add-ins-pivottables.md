---
title: Trabalhar com tabelas dinâmicas usando a API do JavaScript Excel
description: Use a API do JavaScript Excel para criar tabelas dinâmicas e interagir com seus componentes.
ms.date: 09/21/2018
ms.openlocfilehash: a3ff624f8e4e6652834f0a424b482b372c6f2401
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505906"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="07f25-103">Trabalhar com tabelas dinâmicas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="07f25-103">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="07f25-p101">As tabelas dinâmicas simplificam os conjuntos de dados maiores. Permitem a manipulação rápida de dados agrupados. A API JavaScript do Excel possibilita que os suplementos criem tabelas dinâmicas e interajam com seus componentes.</span><span class="sxs-lookup"><span data-stu-id="07f25-p101">PivotTables streamline larger data sets. They allow the quick manipulation of grouped data. The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> 

<span data-ttu-id="07f25-p102">Se não está familiarizado com a funcionalidade das tabelas dinâmicas, considere explorá-las como usuário final. Consulte [Criar uma tabela dinâmica para analisar dados de planilhas](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) para obter uma boa orientação sobre essas ferramentas.</span><span class="sxs-lookup"><span data-stu-id="07f25-p102">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user. See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="07f25-p103">Este artigo fornece exemplos de código para cenários comuns. Para enriquecer a compreensão da API de tabela dinâmica, consulte [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) e [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span><span class="sxs-lookup"><span data-stu-id="07f25-p103">This article provides code samples for common scenarios. To further your understanding of the PivotTable API, see [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="07f25-111">As tabelas dinâmicas criadas com OLAP não são suportadas no momento.</span><span class="sxs-lookup"><span data-stu-id="07f25-111">PivotTables created with OLAP are not currently supported.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="07f25-112">Hierarquias</span><span class="sxs-lookup"><span data-stu-id="07f25-112">Hierarchies</span></span>

<span data-ttu-id="07f25-p104">As tabelas dinâmicas são organizadas com base em quatro categorias de hierarquia: linha, coluna, dados e filtro. Os dados a seguir, que descrevem as vendas de frutas de várias fazendas, serão utilizados ao longo deste artigo.</span><span class="sxs-lookup"><span data-stu-id="07f25-p104">PivotTables are organized based on four hierarchy categories: row, column, data, and filter. The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![Conjunto das vendas de fruta de diferentes tipos provenientes de várias fazendas.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="07f25-p105">Esses dados têm cinco hierarquias: **Fazendas**, **Tipo**, **Classificação**, **Caixas vendidas na fazenda**, e **Caixas vendidas por atacado**. Cada hierarquia só pode existir em uma das quatro categorias. Se **Tipo** for adicionado as hierarquias de coluna e depois adicionado as hierarquias de linha, ele permanecerá apenas no último.</span><span class="sxs-lookup"><span data-stu-id="07f25-p105">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**. Each hierarchy can only exist in one of the four categories. If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="07f25-p106">As hierarquias de linha e coluna definem como os dados serão agrupados. Por exemplo, uma hierarquia de linha de **Fazendas** agrupará todos os conjuntos de dados da mesma fazenda. A escolha entre hierarquia de linha e coluna define a orientação da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="07f25-p106">Row and column hierarchies define how data will be grouped. For example, a row hierarchy of **Farms** will group together all the data sets from the same farm. The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="07f25-p107">As hierarquias de dados são os valores a serem agregados com base nas hierarquias de linhas e colunas. Uma tabela dinâmica com a hierarquia de linhas **Fazendas** e a hierarquia de dados  **Caixas vendidas por atacado** mostra a soma total (por padrão) de todas as frutas diferentes para cada fazenda.</span><span class="sxs-lookup"><span data-stu-id="07f25-p107">Data hierarchies are the values to be aggregated based on the row and column hierarchies. A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="07f25-p108">As hierarquias de filtro incluem ou excluem dados do pivô com base nos valores desse tipo filtrado. Uma hierarquia de filtro de **Classificação** com o tipo **Orgânico** selecionado mostra apenas os dados para fruta orgânica.</span><span class="sxs-lookup"><span data-stu-id="07f25-p108">Filter hierarchies include or exclude data from the pivot based on values within that filtered type. A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="07f25-p109">Aqui estão os dados da fazenda novamente, junto com uma tabela dinâmica. A tabela dinâmica está usando **Fazenda** e **Tipo** como as hierarquias de linha, **Caixas vendidas na fazenda** e **Caixas vendidas por atacado** como as hierarquias de dados (com a função de agregação de soma padrão) e **Classificação** como uma hierarquia de filtro (com **Orgânico** selecionado).</span><span class="sxs-lookup"><span data-stu-id="07f25-p109">Here is the farm data again, alongside a PivotTable. The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![Uma seleção de dados de vendas de frutas ao lado de uma tabela dinâmica com hierarquias de linhas, dados e filtros.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="07f25-p110">Esta tabela dinâmica pode ser gerada por meio da API do JavaScript ou da interface gráfica do Excel. Ambas as opções permitem mais manipulação através de suplementos.</span><span class="sxs-lookup"><span data-stu-id="07f25-p110">This PivotTable could be generated through the JavaScript API or through the Excel UI. Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="07f25-131">Criar uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="07f25-131">Create a PivotTable with Range objects</span></span>

<span data-ttu-id="07f25-p111">Tabelas dinâmicas precisam de um nome, origem e destino. A origem pode ser um endereço de intervalo ou um nome de tabela  (transmitido como um tipo `Range`, `string` ou `Table` ). O destino é um endereço de intervalo (fornecido como `Range` ou `string`). Os exemplos a seguir mostram várias técnicas de criação de tabelas dinâmicas.</span><span class="sxs-lookup"><span data-stu-id="07f25-p111">PivotTables need a name, source, and destination. The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type). The destination is a range address (given as either a `Range` or `string`). The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="07f25-136">Criar uma tabela dinâmica com endereços de intervalo</span><span class="sxs-lookup"><span data-stu-id="07f25-136">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="07f25-137">Criar uma tabela dinâmica com objetos de intervalo</span><span class="sxs-lookup"><span data-stu-id="07f25-137">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="07f25-138">Criar uma tabela dinâmica no nível da pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="07f25-138">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="07f25-139">Usar uma tabela dinâmica existente</span><span class="sxs-lookup"><span data-stu-id="07f25-139">Use an existing PivotTable</span></span>

<span data-ttu-id="07f25-140">As tabelas dinâmicas criadas manualmente também são acessíveis através da coleção de tabela dinâmica da pasta de trabalho ou de planilhas individuais.</span><span class="sxs-lookup"><span data-stu-id="07f25-140">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="07f25-p112">O código a seguir obtém a primeira tabela dinâmica na pasta de trabalho. Em seguida, fornece um nome para a tabela para facilitar a referência futura.</span><span class="sxs-lookup"><span data-stu-id="07f25-p112">The following code gets the first PivotTable in the workbook. It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="07f25-143">Adicionar linhas e colunas à tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="07f25-143">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="07f25-144">As linhas e colunas articulam os dados em torno dos valores desses campos.</span><span class="sxs-lookup"><span data-stu-id="07f25-144">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="07f25-p113">Adicionar a coluna **Fazenda** articula todas as vendas ao redor de cada fazenda. Adicionar as linhas **Tipo** e **Classificação** divide ainda mais os dados com base no tipo de fruta vendida e se a mesma era orgânica ou não.</span><span class="sxs-lookup"><span data-stu-id="07f25-p113">Adding the **Farm** column pivots all the sales around each farm. Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="07f25-148">Você também pode ter uma tabela dinâmica apenas com linhas ou colunas.</span><span class="sxs-lookup"><span data-stu-id="07f25-148">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="07f25-149">Adicionar hierarquias de dados à tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="07f25-149">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="07f25-p114">As hierarquias de dados preenchem a tabela dinâmica com informações para combinar com base nas linhas e colunas. Adicionar as hierarquias de dados de **Caixas vendidas na fazenda** e **Caixas vendidas por atacado** fornece a soma desses números para cada linha e coluna.</span><span class="sxs-lookup"><span data-stu-id="07f25-p114">Data hierarchies fill the PivotTable with information to combine based on the rows and columns. Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="07f25-152">No exemplo, **Fazenda** e **Tipo** são linhas com os dados das vendas de caixas.</span><span class="sxs-lookup"><span data-stu-id="07f25-152">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

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

## <a name="change-aggregation-function"></a><span data-ttu-id="07f25-154">Alterar a função de agregação</span><span class="sxs-lookup"><span data-stu-id="07f25-154">Change aggregation function</span></span>

<span data-ttu-id="07f25-p115">As hierarquias de dados têm seus valores agregados. Para conjuntos de dados de números, por padrão, isso corresponde a uma soma. A propriedade `summarizeBy` define esse comportamento baseando-se em um tipo [AggregrationFunction](https://docs.microsoft.com/javascript/api/excel/excel.aggregationfunction).</span><span class="sxs-lookup"><span data-stu-id="07f25-p115">Data hierarchies have their values aggregated. For datasets of numbers, this is a sum by default. The `summarizeBy` property defines this behavior based on an [](https://docs.microsoft.com/javascript/api/excel/excel.aggregationfunction) type.</span></span> 

<span data-ttu-id="07f25-158">Os tipos de função agregada suportados atualmente são `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP` e `Automatic` (padrão).</span><span class="sxs-lookup"><span data-stu-id="07f25-158">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="07f25-159">O exemplo de código a seguir altera a agregação para as médias dos dados.</span><span class="sxs-lookup"><span data-stu-id="07f25-159">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="07f25-160">Altere os cálculos com ShowAsRule</span><span class="sxs-lookup"><span data-stu-id="07f25-160">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="07f25-p116">As Tabelas Dinâmicas, por padrão, agregam os dados de suas hierarquias de linha e coluna de forma independente. O [ShowAsRule](https://docs.microsoft.com/javascript/api/excel/excel.showasrule) altera a hierarquia dos dados para valores de saída com base em outros itens na tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="07f25-p116">PivotTables, by default, aggregate the data of their row and column hierarchies independently. A [](https://docs.microsoft.com/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="07f25-163">O objeto `ShowAsRule` tem três propriedades:</span><span class="sxs-lookup"><span data-stu-id="07f25-163">The `ShowAsRule` object has three properties:</span></span>
-   <span data-ttu-id="07f25-164">`calculation`: o tipo de cálculo relativo a ser aplicado à hierarquia de dados (o padrão é `none`).</span><span class="sxs-lookup"><span data-stu-id="07f25-164">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
-   <span data-ttu-id="07f25-p117">`baseField`: o campo dentro da hierarquia que contém os dados de base antes que o cálculo seja aplicado. O [PivotField](https://docs.microsoft.com/javascript/api/excel/excel.pivotfield)  geralmente tem o mesmo nome que sua hierarquia pai.</span><span class="sxs-lookup"><span data-stu-id="07f25-p117">`baseField`: The field within the hierarchy containing the base data before the calculation is applied. The [](https://docs.microsoft.com/javascript/api/excel/excel.pivotfield) usually has the same name as its parent hierarchy.</span></span>
-   <span data-ttu-id="07f25-p118">`baseItem`: O item individual [PivotItem](https://docs.microsoft.com/javascript/api/excel/excel.pivotitem) comparado com os valores dos campos de base de acordo com o tipo de cálculo. Nem todos os cálculos exigem esse campo.</span><span class="sxs-lookup"><span data-stu-id="07f25-p118">: The individual item compared against the values of the base fields based on the calculation type. Not all calculations require this field.</span></span>

<span data-ttu-id="07f25-p119">O exemplo a seguir define o cálculo na hierarquia de dados **Soma das caixas vendidas na Fazenda** para uma porcentagem do total da coluna. Ainda queremos que a granularidade se estenda ao nível do tipo de fruta, então usaremos a hierarquia de linha **Tipo** e o campo subjacente. O exemplo também tem **Fazenda** como a primeira hierarquia de linha, de modo que a entrada total da fazenda exibe também a porcentagem que cada fazenda é responsável por produzir.</span><span class="sxs-lookup"><span data-stu-id="07f25-p119">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total. We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field. The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

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

<span data-ttu-id="07f25-p120">O exemplo anterior definiu o cálculo para a coluna, relativo a uma hierarquia de linha individual. Quando o cálculo está relacionado a um item individual, use a propriedade `baseItem` .</span><span class="sxs-lookup"><span data-stu-id="07f25-p120">The previous example set the calculation to the column, relative to an individual row hierarchy. When the calculation relates to an individual item, use the `baseItem` property.</span></span> 

<span data-ttu-id="07f25-p121">O exemplo a seguir mostra o cálculo `differenceFrom` . Exibe a diferença das entradas da hierarquia de dados de vendas de caixas na fazenda em relação  àquelas das "Fazendas A". O `baseField` é **Fazenda**, portanto, vemos as diferenças entre as outras fazendas, bem como as divisões para cada tipo de fruta (**Tipo** também é uma hierarquia de linha neste exemplo).</span><span class="sxs-lookup"><span data-stu-id="07f25-p121">The following example shows the `differenceFrom` calculation. It displays the difference of the farm crate sales data hierarchy entries relative to those of “A Farms”. The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![Uma tabela dinâmica mostrando as diferenças de vendas de frutas entre “Fazendas A” e as outras. Isso mostra a diferença no total de vendas de frutas das fazendas e as vendas de tipos de frutas. Se “Fazendas A” não vendeu um tipo específico de fruta,  é exibida a mensagem “#N/A”.](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="pivottable-layouts"></a><span data-ttu-id="07f25-181">Layouts de tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="07f25-181">PivotTable layouts</span></span>

<span data-ttu-id="07f25-p123">Um [PivotLayout](https://docs.microsoft.com/javascript/api/excel/excel.pivotlayout)  define o posicionamento de hierarquias e seus dados. Você acessa o layout para determinar os intervalos em que os dados são armazenados.</span><span class="sxs-lookup"><span data-stu-id="07f25-p123">A PivotTable layout defines the placement of hierarchies and their data. You access the layout to determine the ranges where data is stored.</span></span> 

<span data-ttu-id="07f25-184">O diagrama a seguir mostra as chamadas de funções de layout que correspondem a cada intervalo da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="07f25-184">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Um diagrama que mostra quais seções de uma tabela dinâmica são retornadas pelas funções de intervalo do layout.](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="07f25-p124">O código a seguir demonstra como obter a última linha dos dados de tabela dinâmica percorrendo o layout. Esses valores são então somados para obter um total geral.</span><span class="sxs-lookup"><span data-stu-id="07f25-p124">The following code demonstrates how to get the last row of the PivotTable data by going through the layout. Those values are then summed together for a grand total.</span></span>

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

<span data-ttu-id="07f25-p125">As tabelas dinâmicas tês três estilos de layout: Compacto, Estrutura do Código e Tabular. Nos exemplos anteriores foi usado o estilo compacto.</span><span class="sxs-lookup"><span data-stu-id="07f25-p125">PivotTables have three layout styles: Compact, Outline, and Tabular. We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="07f25-p126">Os exemplos a seguir usam os estilos de estrutura de código e tabular, respectivamente. O exemplo de código mostra como alternar entre os diferentes layouts.</span><span class="sxs-lookup"><span data-stu-id="07f25-p126">The following examples use the outline and tabular styles, respectively. The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="07f25-192">Layout de estrutura do código</span><span class="sxs-lookup"><span data-stu-id="07f25-192">Outline layout</span></span>

![Uma tabela dinâmica usando o layout de estrutura de tópicos.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="07f25-194">Layout tabular</span><span class="sxs-lookup"><span data-stu-id="07f25-194">Tabular layout</span></span>

![Uma tabela dinâmica usando o layout tabular.](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="07f25-196">Alterar os nomes de hierarquia</span><span class="sxs-lookup"><span data-stu-id="07f25-196">Change hierarchy names</span></span>

<span data-ttu-id="07f25-p127">Os campos de hierarquia são editáveis. O código a seguir demonstra como alterar os nomes exibidos de duas hierarquias de dados.</span><span class="sxs-lookup"><span data-stu-id="07f25-p127">Hierarchy fields are editable. The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="07f25-199">Excluir uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="07f25-199">Delete a PivotTable</span></span>

<span data-ttu-id="07f25-200">As tabelas dinâmicas são excluídas pelo uso de seu nome.</span><span class="sxs-lookup"><span data-stu-id="07f25-200">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="07f25-201">Confira também</span><span class="sxs-lookup"><span data-stu-id="07f25-201">See also</span></span>

- [<span data-ttu-id="07f25-202">Conceitos básicos de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="07f25-202">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="07f25-203">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="07f25-203">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/api/excel)
