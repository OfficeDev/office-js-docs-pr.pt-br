---
title: Trabalhar com tabelas dinâmicas usando a API do JavaScript Excel
description: Use a API do JavaScript Excel para criar tabelas dinâmicas e interagir com seus componentes.
ms.date: 09/21/2018
ms.openlocfilehash: b8704389ced3686858f488b2a50f80c22b1b8bd6
ms.sourcegitcommit: e7e4d08569a01c69168bb005188e9a1e628304b9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/22/2018
ms.locfileid: "24967666"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="c6746-103">Trabalhar com tabelas dinâmicas usando a API do JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="c6746-103">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="c6746-104">As tabelas dinâmicas simplificam conjuntos de dados maiores.</span><span class="sxs-lookup"><span data-stu-id="c6746-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="c6746-105">Elas permitem a rápida manipulação de dados agrupados.</span><span class="sxs-lookup"><span data-stu-id="c6746-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="c6746-106">A API do JavaScript Excel permite que seu suplemento criar tabelas dinâmicas e interaja com seus componentes.</span><span class="sxs-lookup"><span data-stu-id="c6746-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> 

<span data-ttu-id="c6746-107">Se não estiver familiarizado com a funcionalidade das tabelas dinâmicas, considere explorá-las como um usuário final.</span><span class="sxs-lookup"><span data-stu-id="c6746-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end-user.</span></span> <span data-ttu-id="c6746-108">Veja [Criar uma tabela dinâmica para analisar dados de planilha](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) para uma boa orientação sobre essas ferramentas.</span><span class="sxs-lookup"><span data-stu-id="c6746-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="c6746-109">Este artigo fornece exemplos de código para cenários comuns.</span><span class="sxs-lookup"><span data-stu-id="c6746-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="c6746-110">Para enriquecer a compreensão da API de tabela dinâmica, veja [**Tabela dinâmica**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) e [**Coleção de tabelas dinâmicas**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span><span class="sxs-lookup"><span data-stu-id="c6746-110">To further your understanding of the PivotTable API, see [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c6746-111">Tabelas dinâmicas criadas com OLAP não são suportadas no momento.</span><span class="sxs-lookup"><span data-stu-id="c6746-111">PivotTables created with OLAP are not currently supported.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="c6746-112">Hierarquias</span><span class="sxs-lookup"><span data-stu-id="c6746-112">Hierarchies</span></span>

<span data-ttu-id="c6746-113">Tabelas dinâmicas são organizadas com base em quatro categorias de hierarquia: linha, coluna, dados e filtro.</span><span class="sxs-lookup"><span data-stu-id="c6746-113">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="c6746-114">Os dados a seguir que descrevem vendas fruta de diversas fazendas serão usados ao longo deste artigo.</span><span class="sxs-lookup"><span data-stu-id="c6746-114">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![Uma coleção de vendas de frutas de diferentes tipos de diferentes fazendas.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="c6746-116">Esses dados têm cinco hierarquias: **Fazendas**, **Tipo**, **Classificação**, **Caixas vendidas na fazenda**e **Caixas vendidas no atacado**.</span><span class="sxs-lookup"><span data-stu-id="c6746-116">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="c6746-117">Cada hierarquia só pode existir em uma das quatro categorias.</span><span class="sxs-lookup"><span data-stu-id="c6746-117">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="c6746-118">Se **Tipo** for adicionado às hierarquias de coluna e depois for adicionados às hierarquias de linha, ele apenas permanecerá nas últimas.</span><span class="sxs-lookup"><span data-stu-id="c6746-118">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="c6746-119">Hierarquias de linha e de coluna definem como os dados serão agrupados.</span><span class="sxs-lookup"><span data-stu-id="c6746-119">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="c6746-120">Por exemplo, uma hierarquia de linha de **Fazendas** agrupará todos os conjuntos de dados da mesma fazenda.</span><span class="sxs-lookup"><span data-stu-id="c6746-120">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="c6746-121">A escolha entre a hierarquia de linha e de coluna define a orientação da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="c6746-121">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="c6746-122">As hierarquias de dados são os valores a serem agregados com base nas hierarquias de linha e de coluna.</span><span class="sxs-lookup"><span data-stu-id="c6746-122">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="c6746-123">Uma tabela dinâmica com uma hierarquia de linha de **Fazendas** e uma hierarquia de dados de **Caixas vendidas no atacado** mostra a soma total (por padrão) de todas as frutas diferentes para cada fazenda.</span><span class="sxs-lookup"><span data-stu-id="c6746-123">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="c6746-124">Hierarquias de filtro incluem ou excluem dados da tabela dinâmica com base nos valores dentro desse tipo filtrado.</span><span class="sxs-lookup"><span data-stu-id="c6746-124">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="c6746-125">Uma hierarquia de filtro de **Classificação** com o tipo **Orgânico** selecionado mostra apenas os dados de frutas orgânicas.</span><span class="sxs-lookup"><span data-stu-id="c6746-125">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="c6746-126">Aqui estão os dados da fazenda novamente, junto com uma tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="c6746-126">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="c6746-127">A tabela dinâmica está usando **Fazenda** e **Tipo** como as hierarquias de linha, **Caixas vendidas na fazenda** e **Caixas vendidas no atacado** como as hierarquias de dados (com a função de agregação padrão de soma) e **Classificação** como uma hierarquia de filtro (com **Orgânico** selecionado).</span><span class="sxs-lookup"><span data-stu-id="c6746-127">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![Uma seleção de dados de vendas de frutas ao lado de uma tabela dinâmica com hierarquias de linha, dados e filtro.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="c6746-129">Esta tabela dinâmica poderia ser gerada por meio da API do JavaScript ou por meio da interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="c6746-129">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="c6746-130">As duas opções permitem manipulação adicional por meio de suplementos.</span><span class="sxs-lookup"><span data-stu-id="c6746-130">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="c6746-131">Criar uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="c6746-131">Create a PivotTable with Range objects</span></span>

<span data-ttu-id="c6746-132">Tabelas dinâmicas precisam de um nome, origem e destino.</span><span class="sxs-lookup"><span data-stu-id="c6746-132">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="c6746-133">A fonte pode ser um endereço de intervalo ou um nome da tabela (passado como um tipo `Range`, `string` ou `Table`).</span><span class="sxs-lookup"><span data-stu-id="c6746-133">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="c6746-134">O destino é um endereço de intervalo (dado como um `Range` ou `string`).</span><span class="sxs-lookup"><span data-stu-id="c6746-134">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="c6746-135">Os exemplos a seguir mostram várias técnicas de criação de uma tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="c6746-135">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="c6746-136">Criar uma tabela dinâmica com o endereços de intervalo</span><span class="sxs-lookup"><span data-stu-id="c6746-136">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" created on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="c6746-137">Criar uma tabela dinâmica com objetos de intervalo</span><span class="sxs-lookup"><span data-stu-id="c6746-137">Create a PivotTable with Range objects</span></span>

```typescript
await Excel.run(async (context) => {    
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21
    const rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    const rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);
    
    await context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="c6746-138">Criar uma tabela dinâmica no nível da pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="c6746-138">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="c6746-139">Usar uma tabela dinâmica existente</span><span class="sxs-lookup"><span data-stu-id="c6746-139">Use an existing PivotTable</span></span>

<span data-ttu-id="c6746-140">A criação de tabelas dinâmicas manualmente também é acessível por meio da coleção de tabelas dinâmicas da pasta de trabalho ou das planilhas individuais.</span><span class="sxs-lookup"><span data-stu-id="c6746-140">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="c6746-141">O código a seguir obtém a primeira tabela dinâmica na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="c6746-141">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="c6746-142">Ele então oferece um nome para a tabela para facilitar a referência futura.</span><span class="sxs-lookup"><span data-stu-id="c6746-142">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="c6746-143">Adicionar linhas e colunas à tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="c6746-143">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="c6746-144">Linhas e colunas articulam os dados em torno desses campos de valores.</span><span class="sxs-lookup"><span data-stu-id="c6746-144">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="c6746-145">Adicionar a coluna **Fazenda** articula todas as vendas ao redor de cada fazenda.</span><span class="sxs-lookup"><span data-stu-id="c6746-145">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="c6746-146">Adicionar as linhas **Tipo** e **Classificação** quebra os dados com base em qual fruta foi vendida e se ela era orgânica ou não.</span><span class="sxs-lookup"><span data-stu-id="c6746-146">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="c6746-148">Você também pode ter uma tabela dinâmica apenas com linhas ou colunas.</span><span class="sxs-lookup"><span data-stu-id="c6746-148">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="c6746-149">Adicionar hierarquias de dados à tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="c6746-149">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="c6746-150">Hierarquias de dados preenchem a tabela dinâmica com informações para combinar com base nas linhas e colunas.</span><span class="sxs-lookup"><span data-stu-id="c6746-150">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="c6746-151">Adicionar as hierarquias de dados de **Caixas vendidas na fazenda** e **Caixas vendidas no atacado** dá somas àqueles valores para cada linha e coluna.</span><span class="sxs-lookup"><span data-stu-id="c6746-151">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="c6746-152">No exemplo, tanto **Fazenda** quanto **Tipo** são linhas, com as vendas das caixas como os dados.</span><span class="sxs-lookup"><span data-stu-id="c6746-152">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

![Uma tabela dinâmica que mostra as vendas totais das diferentes frutas com base na fazenda de onde elas vieram.](../images/excel-pivots-data-hierarchy.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the heirarchies that will have their data aggregated (summed in this case)
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    await context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="c6746-154">Alterar a função de agregação</span><span class="sxs-lookup"><span data-stu-id="c6746-154">Change aggregation function</span></span>

<span data-ttu-id="c6746-155">Hierarquias de dados têm seus valores agregados.</span><span class="sxs-lookup"><span data-stu-id="c6746-155">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="c6746-156">Para conjuntos de dados de números, essa é uma soma por padrão.</span><span class="sxs-lookup"><span data-stu-id="c6746-156">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="c6746-157">A propriedade `summarizeBy` define esse comportamento com base em um tipo `AggregrationFunction`.</span><span class="sxs-lookup"><span data-stu-id="c6746-157">The `summarizeBy` property defines this behavior based on an `AggregrationFunction` type.</span></span> 

<span data-ttu-id="c6746-158">Os tipos de função agregada suportados atualmente são `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP` e `Automatic` (padrão).</span><span class="sxs-lookup"><span data-stu-id="c6746-158">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="c6746-159">Os exemplos de códigos a seguir alteram a agregação para as médias dos dados.</span><span class="sxs-lookup"><span data-stu-id="c6746-159">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="pivottable-layouts"></a><span data-ttu-id="c6746-160">Layouts de tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="c6746-160">PivotTable layouts</span></span>

<span data-ttu-id="c6746-161">Um layout de tabela dinâmica define o posicionamento das hierarquias e seus dados.</span><span class="sxs-lookup"><span data-stu-id="c6746-161">A PivotTable layout defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="c6746-162">Você acessa o layout para determinar os intervalos de onde os dados são armazenados.</span><span class="sxs-lookup"><span data-stu-id="c6746-162">You access the layout to determine the ranges where data is stored.</span></span> 

<span data-ttu-id="c6746-163">O diagrama a seguir mostra qual chamadas de funções de layout correspondem a quais intervalos da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="c6746-163">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Um diagrama que mostra quais seções de uma tabela dinâmica são retornadas pelas funções de intervalo obtidas do layout.](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="c6746-165">O código a seguir demonstra como obter a última linha de dados da tabela dinâmica por meio do layout.</span><span class="sxs-lookup"><span data-stu-id="c6746-165">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="c6746-166">Esses valores são somados para um total geral.</span><span class="sxs-lookup"><span data-stu-id="c6746-166">Those values are then summed together for a grand total.</span></span>


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

<span data-ttu-id="c6746-167">As tabelas dinâmicas têm três estilos de layout: Compacto, Contorno e Tabular.</span><span class="sxs-lookup"><span data-stu-id="c6746-167">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="c6746-168">Vimos o estilo compacto nos exemplos anteriores.</span><span class="sxs-lookup"><span data-stu-id="c6746-168">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="c6746-169">Os exemplos a seguir usam os estilos contorno e tabular, respectivamente.</span><span class="sxs-lookup"><span data-stu-id="c6746-169">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="c6746-170">O exemplo de código mostra como fazer o ciclo entre os diferentes layouts.</span><span class="sxs-lookup"><span data-stu-id="c6746-170">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="c6746-171">Layout contorno</span><span class="sxs-lookup"><span data-stu-id="c6746-171">Outline layout</span></span>

![Uma tabela dinâmica usando o layout de estrutura de tópicos.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="c6746-173">Layout tabular</span><span class="sxs-lookup"><span data-stu-id="c6746-173">Tabular layout</span></span>

![Uma tabela dinâmica usando o layout tabular.](../images/excel-pivots-tabular-layout.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.layout.load("layoutType");
    await context.sync();
    
    // cycling through layout styles
    if (pivotTable.layout.layoutType === "Compact") {
        pivotTable.layout.layoutType = "Outline";
    } else if (pivotTable.layout.layoutType === "Outline") {
        pivotTable.layout.layoutType = "Tabular";
    } else {
        pivotTable.layout.layoutType = "Compact";
    }
    
    await context.sync();
});
```

## <a name="change-hierarchy-names"></a><span data-ttu-id="c6746-175">Alterar os nomes de hierarquia</span><span class="sxs-lookup"><span data-stu-id="c6746-175">Change hierarchy names</span></span>

<span data-ttu-id="c6746-176">Campos de hierarquia são editáveis.</span><span class="sxs-lookup"><span data-stu-id="c6746-176">Hierarchy fields are editable.</span></span> <span data-ttu-id="c6746-177">O código a seguir demonstra como alterar os nomes exibidos de duas hierarquias de dados.</span><span class="sxs-lookup"><span data-stu-id="c6746-177">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="c6746-178">Excluir uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="c6746-178">Delete a PivotTable</span></span>

<span data-ttu-id="c6746-179">As tabelas dinâmicas são excluídas pelo uso de seu nome.</span><span class="sxs-lookup"><span data-stu-id="c6746-179">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="c6746-180">Confira também</span><span class="sxs-lookup"><span data-stu-id="c6746-180">See also</span></span>

- [<span data-ttu-id="c6746-181">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="c6746-181">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c6746-182">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="c6746-182">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/api/excel)
