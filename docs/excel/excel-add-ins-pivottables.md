---
title: Trabalhar com tabelas dinâmicas usando a API do JavaScript Excel
description: Use a API do JavaScript Excel para criar tabelas dinâmicas e interagir com seus componentes.
ms.date: 09/21/2018
ms.openlocfilehash: 5245665bad2933df205bcda29e226a965de1c356
ms.sourcegitcommit: 64da9ed76d22b14df745b1f0ef97a8f5194400e4
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/03/2018
ms.locfileid: "25361021"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="ab73b-103">Trabalhar com tabelas dinâmicas usando a API do JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="ab73b-103">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="ab73b-104">As tabelas dinâmicas simplificam conjuntos de dados maiores.</span><span class="sxs-lookup"><span data-stu-id="ab73b-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="ab73b-105">Elas permitem a rápida manipulação de dados agrupados.</span><span class="sxs-lookup"><span data-stu-id="ab73b-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="ab73b-106">A API JavaScript do Excel permite que seu suplemento crie tabelas dinâmicas e interaja com os seus componentes.</span><span class="sxs-lookup"><span data-stu-id="ab73b-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> 

<span data-ttu-id="ab73b-107">Se você não estiver familiarizado com a funcionalidade de tabelas dinâmicas, considere explorá-las como um usuário final.</span><span class="sxs-lookup"><span data-stu-id="ab73b-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end-user.</span></span> <span data-ttu-id="ab73b-108">Confira [Criar uma tabela dinâmica para analisar dados de planilha](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) para uma boa introdução sobre essas ferramentas.</span><span class="sxs-lookup"><span data-stu-id="ab73b-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="ab73b-109">Este artigo fornece exemplos de código para cenários comuns.</span><span class="sxs-lookup"><span data-stu-id="ab73b-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="ab73b-110">Para enriquecer a compreensão da API de tabela dinâmica, veja [**Tabela dinâmica**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) e [**Coleção de tabelas dinâmicas**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span><span class="sxs-lookup"><span data-stu-id="ab73b-110">To further your understanding of the PivotTable API, see [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ab73b-111">Tabelas dinâmicas criadas com OLAP não são suportadas no momento.</span><span class="sxs-lookup"><span data-stu-id="ab73b-111">PivotTables created with OLAP are not currently supported.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="ab73b-112">Hierarquias</span><span class="sxs-lookup"><span data-stu-id="ab73b-112">Hierarchies</span></span>

<span data-ttu-id="ab73b-113">Tabelas dinâmicas são organizadas com base em quatro categorias de hierarquia: linha, coluna, dados e filtro.</span><span class="sxs-lookup"><span data-stu-id="ab73b-113">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="ab73b-114">Os dados a seguir que descrevem vendas fruta de diversas fazendas serão usados ao longo deste artigo.</span><span class="sxs-lookup"><span data-stu-id="ab73b-114">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![Uma coleção de vendas de frutas de diferentes tipos de diferentes fazendas.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="ab73b-116">Esses dados têm cinco hierarquias: **Fazendas**, **Tipo**, **Classificação**, **Caixas vendidas na fazenda**e **Caixas vendidas no atacado**.</span><span class="sxs-lookup"><span data-stu-id="ab73b-116">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="ab73b-117">Cada hierarquia só pode existir em uma das quatro categorias.</span><span class="sxs-lookup"><span data-stu-id="ab73b-117">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="ab73b-118">Se **Tipo** for adicionado às hierarquias de coluna e depois for adicionados às hierarquias de linha, ele apenas permanecerá nas últimas.</span><span class="sxs-lookup"><span data-stu-id="ab73b-118">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="ab73b-119">Hierarquias de linha e de coluna definem como os dados serão agrupados.</span><span class="sxs-lookup"><span data-stu-id="ab73b-119">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="ab73b-120">Por exemplo, uma hierarquia de linha de **Fazendas** agrupará todos os conjuntos de dados da mesma fazenda.</span><span class="sxs-lookup"><span data-stu-id="ab73b-120">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="ab73b-121">A escolha entre a hierarquia de linha e de coluna define a orientação da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="ab73b-121">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="ab73b-122">As hierarquias de dados são os valores a serem agregados com base nas hierarquias de linha e de coluna.</span><span class="sxs-lookup"><span data-stu-id="ab73b-122">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="ab73b-123">Uma tabela dinâmica com uma hierarquia de linha de **Fazendas** e uma hierarquia de dados de **Caixas vendidas no atacado** mostra a soma total (por padrão) de todas as frutas diferentes para cada fazenda.</span><span class="sxs-lookup"><span data-stu-id="ab73b-123">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="ab73b-124">Hierarquias de filtro incluem ou excluem dados da tabela dinâmica com base nos valores dentro desse tipo filtrado.</span><span class="sxs-lookup"><span data-stu-id="ab73b-124">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="ab73b-125">Uma hierarquia de filtro de **Classificação** com o tipo **Orgânico** selecionado mostra apenas os dados de frutas orgânicas.</span><span class="sxs-lookup"><span data-stu-id="ab73b-125">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="ab73b-126">Aqui estão os dados da fazenda novamente, junto com uma tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="ab73b-126">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="ab73b-127">A tabela dinâmica está usando **Fazenda** e **Tipo** como as hierarquias de linha, **Caixas vendidas na fazenda** e **Caixas vendidas no atacado** como as hierarquias de dados (com a função de agregação padrão de soma) e **Classificação** como uma hierarquia de filtro (com **Orgânico** selecionado).</span><span class="sxs-lookup"><span data-stu-id="ab73b-127">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![Uma seleção de dados de vendas de frutas ao lado de uma tabela dinâmica com hierarquias de linha, dados e filtro.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="ab73b-129">Esta tabela dinâmica poderia ser gerada por meio da API do JavaScript ou por meio da interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="ab73b-129">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="ab73b-130">As duas opções permitem manipulação adicional por meio de suplementos.</span><span class="sxs-lookup"><span data-stu-id="ab73b-130">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="ab73b-131">Criar uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="ab73b-131">Create a PivotTable with Range objects</span></span>

<span data-ttu-id="ab73b-132">Tabelas dinâmicas precisam de um nome, origem e destino.</span><span class="sxs-lookup"><span data-stu-id="ab73b-132">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="ab73b-133">A fonte pode ser um endereço de intervalo ou um nome da tabela (passado como um tipo `Range`, `string` ou `Table`).</span><span class="sxs-lookup"><span data-stu-id="ab73b-133">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="ab73b-134">O destino é um endereço de intervalo (dado como um `Range` ou `string`).</span><span class="sxs-lookup"><span data-stu-id="ab73b-134">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="ab73b-135">Os exemplos a seguir mostram várias técnicas de criação de uma tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="ab73b-135">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="ab73b-136">Criar uma tabela dinâmica com o endereços de intervalo</span><span class="sxs-lookup"><span data-stu-id="ab73b-136">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="ab73b-137">Criar uma tabela dinâmica com objetos de intervalo</span><span class="sxs-lookup"><span data-stu-id="ab73b-137">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="ab73b-138">Criar uma tabela dinâmica no nível da pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="ab73b-138">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="ab73b-139">Usar uma tabela dinâmica existente</span><span class="sxs-lookup"><span data-stu-id="ab73b-139">Use an existing PivotTable</span></span>

<span data-ttu-id="ab73b-140">A criação de tabelas dinâmicas manualmente também é acessível por meio da coleção de tabelas dinâmicas da pasta de trabalho ou das planilhas individuais.</span><span class="sxs-lookup"><span data-stu-id="ab73b-140">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="ab73b-141">O código a seguir obtém a primeira tabela dinâmica na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="ab73b-141">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="ab73b-142">Ele então oferece um nome para a tabela para facilitar a referência futura.</span><span class="sxs-lookup"><span data-stu-id="ab73b-142">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="ab73b-143">Adicionar linhas e colunas à tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="ab73b-143">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="ab73b-144">Linhas e colunas articulam os dados em torno desses campos de valores.</span><span class="sxs-lookup"><span data-stu-id="ab73b-144">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="ab73b-145">Adicionar a coluna **Fazenda** articula todas as vendas ao redor de cada fazenda.</span><span class="sxs-lookup"><span data-stu-id="ab73b-145">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="ab73b-146">Adicionar as linhas **Tipo** e **Classificação** quebra os dados com base em qual fruta foi vendida e se ela era orgânica ou não.</span><span class="sxs-lookup"><span data-stu-id="ab73b-146">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

![Uma tabela dinâmica com uma coluna Fazenda e linhas Tipo e Classificação.](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

<span data-ttu-id="ab73b-148">Você também pode ter uma tabela dinâmica apenas com linhas ou colunas.</span><span class="sxs-lookup"><span data-stu-id="ab73b-148">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="ab73b-149">Adicionar hierarquias de dados à tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="ab73b-149">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="ab73b-150">Hierarquias de dados preenchem a tabela dinâmica com informações para combinar com base nas linhas e colunas.</span><span class="sxs-lookup"><span data-stu-id="ab73b-150">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="ab73b-151">Adicionar as hierarquias de dados de **Caixas vendidas na fazenda** e **Caixas vendidas no atacado** dá somas àqueles valores para cada linha e coluna.</span><span class="sxs-lookup"><span data-stu-id="ab73b-151">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="ab73b-152">No exemplo, tanto **Fazenda** quanto **Tipo** são linhas, com as vendas das caixas como os dados.</span><span class="sxs-lookup"><span data-stu-id="ab73b-152">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

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

## <a name="change-aggregation-function"></a><span data-ttu-id="ab73b-154">Alterar a função de agregação</span><span class="sxs-lookup"><span data-stu-id="ab73b-154">Change aggregation function</span></span>

<span data-ttu-id="ab73b-155">Hierarquias de dados têm seus valores agregados.</span><span class="sxs-lookup"><span data-stu-id="ab73b-155">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="ab73b-156">Para conjuntos de dados de números, essa é uma soma por padrão.</span><span class="sxs-lookup"><span data-stu-id="ab73b-156">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="ab73b-157">A propriedade `summarizeBy` define esse comportamento com base em um tipo `AggregrationFunction`.</span><span class="sxs-lookup"><span data-stu-id="ab73b-157">The `summarizeBy` property defines this behavior based on an `AggregrationFunction` type.</span></span> 

<span data-ttu-id="ab73b-158">Os tipos de função agregada suportados atualmente são `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP` e `Automatic` (padrão).</span><span class="sxs-lookup"><span data-stu-id="ab73b-158">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="ab73b-159">O exemplo de código a seguir altera a agregação para as médias dos dados.</span><span class="sxs-lookup"><span data-stu-id="ab73b-159">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="ab73b-160">Altere os cálculos com ShowAsRule</span><span class="sxs-lookup"><span data-stu-id="ab73b-160">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="ab73b-161">As tabelas dinâmicas, por padrão, agregam os dados de suas hierarquias de linha e coluna de forma independente.</span><span class="sxs-lookup"><span data-stu-id="ab73b-161">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="ab73b-162">Uma `ShowAsRule` altera a hierarquia dos dados para valores de saída com base em outros itens da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="ab73b-162">A `ShowAsRule` changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="ab73b-163">O objeto `ShowAsRule` tem três propriedades:</span><span class="sxs-lookup"><span data-stu-id="ab73b-163">The `ShowAsRule` object has three properties:</span></span>
-   <span data-ttu-id="ab73b-164">`calculation`: O tipo de cálculo relativo para aplicar à hierarquia de dados (o padrão é `none`).</span><span class="sxs-lookup"><span data-stu-id="ab73b-164">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
-   <span data-ttu-id="ab73b-165">`baseField`: O campo dentro da hierarquia que contém os dados de base antes do cálculo ser aplicado.</span><span class="sxs-lookup"><span data-stu-id="ab73b-165">`baseField`: The field within the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="ab73b-166">O `PivotField` normalmente tem o mesmo nome que sua hierarquia pai.</span><span class="sxs-lookup"><span data-stu-id="ab73b-166">The `PivotField` usually has the same name as its parent hierarchy.</span></span>
-   <span data-ttu-id="ab73b-167">`baseItem`: O item individual comparado com os valores dos campos de base de acordo com o tipo de cálculo.</span><span class="sxs-lookup"><span data-stu-id="ab73b-167">`baseItem`: The individual item compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="ab73b-168">Nem todos os cálculos exigem esse campo.</span><span class="sxs-lookup"><span data-stu-id="ab73b-168">Not all calculations require this field.</span></span>

<span data-ttu-id="ab73b-169">O exemplo a seguir define o cálculo da hierarquia de dados **Soma de caixas vendidas na fazenda** como uma porcentagem do total de coluna.</span><span class="sxs-lookup"><span data-stu-id="ab73b-169">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span> <span data-ttu-id="ab73b-170">Ainda queremos a granularidade para estender o nível de tipo de fruta, então usaremos a hierarquia de linha de **Tipo** e seu campo subjacente.</span><span class="sxs-lookup"><span data-stu-id="ab73b-170">We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field.</span></span> <span data-ttu-id="ab73b-171">O exemplo também tem **Fazenda** como a primeira linha da hierarquia, assim, o total de entradas da fazenda também mostra a porcentagem que cada fazenda é responsável por produzir.</span><span class="sxs-lookup"><span data-stu-id="ab73b-171">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![Uma tabela dinâmica mostrando as porcentagens de venda de frutas em relação ao total geral, tanto por fazenda quanto por tipo de fruta dentro de cada fazenda.](../images/excel-pivots-showas-percentage.png)

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

<span data-ttu-id="ab73b-173">O exemplo anterior define o cálculo para a coluna em relação à hierarquia de uma linha individual.</span><span class="sxs-lookup"><span data-stu-id="ab73b-173">The previous example set the calculation to the column, relative to an individual row hierarchy.</span></span> <span data-ttu-id="ab73b-174">Quando o cálculo está relacionado a um item individual, use a propriedade `baseItem`.</span><span class="sxs-lookup"><span data-stu-id="ab73b-174">When the calculation relates to an individual item, use the `baseItem` property.</span></span> 

<span data-ttu-id="ab73b-175">O exemplo a seguir mostra o cálculo `differenceFrom`.</span><span class="sxs-lookup"><span data-stu-id="ab73b-175">The following example shows the request.</span></span> <span data-ttu-id="ab73b-176">Ele mostra a diferença entre das entradas na hierarquia de dados de vendas de caixas na fazenda em relação à "Fazendas A".</span><span class="sxs-lookup"><span data-stu-id="ab73b-176">It displays the difference of the farm crate sales data hierarchy entries relative to those of “A Farms”.</span></span> <span data-ttu-id="ab73b-177">O `baseField` é **Fazenda**, portanto, vemos as diferenças entre as outras fazendas, bem como o detalhe de cada tipo de fruta (**Tipo** também é uma hierarquia de linha neste exemplo).</span><span class="sxs-lookup"><span data-stu-id="ab73b-177">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![Uma tabela dinâmica mostrando as diferenças das vendas de fruta entre "Fazendas A" e as outras.](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="pivottable-layouts"></a><span data-ttu-id="ab73b-181">Layouts de tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="ab73b-181">PivotTable layouts</span></span>

<span data-ttu-id="ab73b-182">Um layout de tabela dinâmica define o posicionamento das hierarquias e seus dados.</span><span class="sxs-lookup"><span data-stu-id="ab73b-182">A PivotTable layout defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="ab73b-183">Você acessa o layout para determinar os intervalos de onde os dados são armazenados.</span><span class="sxs-lookup"><span data-stu-id="ab73b-183">You access the layout to determine the ranges where data is stored.</span></span> 

<span data-ttu-id="ab73b-184">O diagrama a seguir mostra qual chamadas de funções de layout correspondem a quais intervalos da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="ab73b-184">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Um diagrama que mostra quais seções de uma tabela dinâmica são retornadas pelas funções de intervalo obtidas do layout.](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="ab73b-186">O código a seguir demonstra como obter a última linha de dados da tabela dinâmica por meio do layout.</span><span class="sxs-lookup"><span data-stu-id="ab73b-186">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="ab73b-187">Esses valores são somados para um total geral.</span><span class="sxs-lookup"><span data-stu-id="ab73b-187">Those values are then summed together for a grand total.</span></span>

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

<span data-ttu-id="ab73b-188">As tabelas dinâmicas têm três estilos de layout: Compacto, Contorno e Tabular.</span><span class="sxs-lookup"><span data-stu-id="ab73b-188">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="ab73b-189">Vimos o estilo compacto nos exemplos anteriores.</span><span class="sxs-lookup"><span data-stu-id="ab73b-189">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="ab73b-190">Os exemplos a seguir usam os estilos contorno e tabular, respectivamente.</span><span class="sxs-lookup"><span data-stu-id="ab73b-190">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="ab73b-191">O exemplo de código mostra como fazer o ciclo entre os diferentes layouts.</span><span class="sxs-lookup"><span data-stu-id="ab73b-191">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="ab73b-192">Layout contorno</span><span class="sxs-lookup"><span data-stu-id="ab73b-192">Outline layout</span></span>

![Uma tabela dinâmica usando o layout de estrutura de tópicos.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="ab73b-194">Layout tabular</span><span class="sxs-lookup"><span data-stu-id="ab73b-194">Tabular layout</span></span>

![Uma tabela dinâmica usando o layout tabular.](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="ab73b-196">Alterar os nomes de hierarquia</span><span class="sxs-lookup"><span data-stu-id="ab73b-196">Change hierarchy names</span></span>

<span data-ttu-id="ab73b-197">Campos de hierarquia são editáveis.</span><span class="sxs-lookup"><span data-stu-id="ab73b-197">Hierarchy fields are editable.</span></span> <span data-ttu-id="ab73b-198">O código a seguir demonstra como alterar os nomes exibidos de duas hierarquias de dados.</span><span class="sxs-lookup"><span data-stu-id="ab73b-198">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="ab73b-199">Excluir uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="ab73b-199">Delete a PivotTable</span></span>

<span data-ttu-id="ab73b-200">As tabelas dinâmicas são excluídas pelo uso de seu nome.</span><span class="sxs-lookup"><span data-stu-id="ab73b-200">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="ab73b-201">Confira também</span><span class="sxs-lookup"><span data-stu-id="ab73b-201">See also</span></span>

- [<span data-ttu-id="ab73b-202">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="ab73b-202">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="ab73b-203">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="ab73b-203">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/api/excel)
