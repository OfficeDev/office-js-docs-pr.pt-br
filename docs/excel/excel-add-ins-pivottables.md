---
title: Trabalhar com tabelas dinâmicas usando a API JavaScript do Excel
description: Use a API JavaScript do Excel para criar tabelas dinâmicas e interagir com seus componentes.
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: b53d734e676417a6438f1008bac720a38a244d1f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449336"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="47e8a-103">Trabalhar com tabelas dinâmicas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="47e8a-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="47e8a-104">As tabelas dinâmicas simplificam conjuntos de dados maiores.</span><span class="sxs-lookup"><span data-stu-id="47e8a-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="47e8a-105">Eles permitem a manipulação rápida dos dados agrupados.</span><span class="sxs-lookup"><span data-stu-id="47e8a-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="47e8a-106">A API JavaScript do Excel permite que o suplemento crie tabelas dinâmicas e interaja com seus componentes.</span><span class="sxs-lookup"><span data-stu-id="47e8a-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span>

<span data-ttu-id="47e8a-107">Se você não estiver familiarizado com a funcionalidade das tabelas dinâmicas, considere explorá-las como um usuário final.</span><span class="sxs-lookup"><span data-stu-id="47e8a-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span> <span data-ttu-id="47e8a-108">ConFira [criar uma tabela dinâmica para analisar os dados da planilha](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) para obter uma boa opção mais interessante nessas ferramentas.</span><span class="sxs-lookup"><span data-stu-id="47e8a-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="47e8a-109">Este artigo fornece exemplos de código para cenários comuns.</span><span class="sxs-lookup"><span data-stu-id="47e8a-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="47e8a-110">Para saber mais sobre a API de tabela dinâmica, confira [**tabela dinâmica**](/javascript/api/excel/excel.pivottable) e [**tabela dinâmica**](/javascript/api/excel/excel.pivottable).</span><span class="sxs-lookup"><span data-stu-id="47e8a-110">To further your understanding of the PivotTable API, see [**PivotTable**](/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="47e8a-111">As tabelas dinâmicas criadas com OLAP não têm suporte no momento.</span><span class="sxs-lookup"><span data-stu-id="47e8a-111">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="47e8a-112">Também não há suporte para o Power pivot.</span><span class="sxs-lookup"><span data-stu-id="47e8a-112">There is also no support for Power Pivot.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="47e8a-113">Hierarquias</span><span class="sxs-lookup"><span data-stu-id="47e8a-113">Hierarchies</span></span>

<span data-ttu-id="47e8a-114">As tabelas dinâmicas são organizadas com base em quatro categorias de hierarquia: linha, coluna, dados e filtro.</span><span class="sxs-lookup"><span data-stu-id="47e8a-114">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="47e8a-115">Os dados a seguir que descrevem as vendas de frutas de vários farms serão usados neste artigo.</span><span class="sxs-lookup"><span data-stu-id="47e8a-115">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![Uma coleção de vendas de frutas de diferentes tipos de farms diferentes.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="47e8a-117">Esses dados têm cinco hierarquias: **farms**, **tipo**, **classificação**, enquando são **vendidas no farm**e as vendidas no **atacado**.</span><span class="sxs-lookup"><span data-stu-id="47e8a-117">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="47e8a-118">Cada hierarquia só pode existir em uma das quatro categorias.</span><span class="sxs-lookup"><span data-stu-id="47e8a-118">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="47e8a-119">Se **Type** for adicionado a hierarquias de coluna e, em seguida, adicionado às hierarquias de linha, ele permanecerá somente no último.</span><span class="sxs-lookup"><span data-stu-id="47e8a-119">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="47e8a-120">Hierarquias de linha e coluna definem como os dados serão agrupados.</span><span class="sxs-lookup"><span data-stu-id="47e8a-120">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="47e8a-121">Por exemplo, uma hierarquia de linha \*\*\*\* de farms agrupará todos os conjuntos de dados do mesmo farm.</span><span class="sxs-lookup"><span data-stu-id="47e8a-121">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="47e8a-122">A escolha entre hierarquia de linha e coluna define a orientação da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="47e8a-122">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="47e8a-123">Hierarquias de dados são os valores a serem agregados com base nas hierarquias de linha e coluna.</span><span class="sxs-lookup"><span data-stu-id="47e8a-123">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="47e8a-124">Uma tabela dinâmica com uma hierarquia de \*\*\*\* linha de farms e uma hierarquia de dados de **envenda vendida** mostra a soma total (por padrão) de todos os diferentes frutas para cada farm.</span><span class="sxs-lookup"><span data-stu-id="47e8a-124">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="47e8a-125">As hierarquias de filtro incluem ou excluem dados da tabela dinâmica com base nos valores desse tipo filtrado.</span><span class="sxs-lookup"><span data-stu-id="47e8a-125">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="47e8a-126">Uma hierarquia de filtro de **classificação** com o tipo **orgânica** selecionado mostra apenas dados para frutas orgânicas.</span><span class="sxs-lookup"><span data-stu-id="47e8a-126">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="47e8a-127">Estes são os dados do farm novamente, juntamente com uma tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="47e8a-127">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="47e8a-128">A tabela dinâmica está usando o **farm** e o **tipo** como hierarquias de linha, as televendedas **no farm** e as doutilizações **vendidas** como as hierarquias de dados (com a função de agregação padrão de Sum) e a **classificação** como um filtro hierarquia (com a **orgânica** selecionada).</span><span class="sxs-lookup"><span data-stu-id="47e8a-128">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![Uma seleção de dados de vendas de frutas ao lado de uma tabela dinâmica com hierarquias de linha, dados e filtros.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="47e8a-130">Esta tabela dinâmica pode ser gerada por meio da API JavaScript ou através da interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="47e8a-130">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="47e8a-131">Ambas as opções permitem mais manipulação por meio de suplementos.</span><span class="sxs-lookup"><span data-stu-id="47e8a-131">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="47e8a-132">Criar uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="47e8a-132">Create a PivotTable</span></span>

<span data-ttu-id="47e8a-133">As tabelas dinâmicas precisam de um nome, origem e destino.</span><span class="sxs-lookup"><span data-stu-id="47e8a-133">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="47e8a-134">A origem pode ser um endereço de intervalo ou nome de tabela (passado `Range`como `string`um, `Table` ou tipo).</span><span class="sxs-lookup"><span data-stu-id="47e8a-134">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="47e8a-135">O destino é um endereço de intervalo (fornecido como ou `Range` um `string`ou).</span><span class="sxs-lookup"><span data-stu-id="47e8a-135">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="47e8a-136">Os exemplos a seguir mostram várias técnicas de criação de tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="47e8a-136">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="47e8a-137">Criar uma tabela dinâmica com endereços de intervalo</span><span class="sxs-lookup"><span data-stu-id="47e8a-137">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="47e8a-138">Criar uma tabela dinâmica com objetos Range</span><span class="sxs-lookup"><span data-stu-id="47e8a-138">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="47e8a-139">Criar uma tabela dinâmica no nível da pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="47e8a-139">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="47e8a-140">Usar uma tabela dinâmica existente</span><span class="sxs-lookup"><span data-stu-id="47e8a-140">Use an existing PivotTable</span></span>

<span data-ttu-id="47e8a-141">As tabelas dinâmicas criadas manualmente também podem ser acessadas por meio da coleção PivotTable da pasta de trabalho ou de planilhas individuais.</span><span class="sxs-lookup"><span data-stu-id="47e8a-141">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="47e8a-142">O código a seguir obtém a primeira tabela dinâmica na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="47e8a-142">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="47e8a-143">Em seguida, ele fornece ao nome da tabela uma referência fácil no futuro.</span><span class="sxs-lookup"><span data-stu-id="47e8a-143">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="47e8a-144">Adicionar linhas e colunas a uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="47e8a-144">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="47e8a-145">Linhas e colunas dinamizam os dados em torno dos valores dos campos.</span><span class="sxs-lookup"><span data-stu-id="47e8a-145">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="47e8a-146">A adição da coluna do **farm** dinamiza todas as vendas em torno de cada farm.</span><span class="sxs-lookup"><span data-stu-id="47e8a-146">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="47e8a-147">Adicionar as linhas de **tipo** e **classificação** divide ainda mais os dados com base no que frutas foi vendido e se foi orgânica ou não.</span><span class="sxs-lookup"><span data-stu-id="47e8a-147">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

![Uma tabela dinâmica com uma coluna do farm e linhas de tipo e classificação.](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

<span data-ttu-id="47e8a-149">Você também pode ter uma tabela dinâmica com apenas linhas ou colunas.</span><span class="sxs-lookup"><span data-stu-id="47e8a-149">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="47e8a-150">Adicionar hierarquias de dados à tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="47e8a-150">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="47e8a-151">As hierarquias de dados preenchem a tabela dinâmica com informações para combinar com base nas linhas e colunas.</span><span class="sxs-lookup"><span data-stu-id="47e8a-151">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="47e8a-152">Adicionar as hierarquias de dados das pessoas **vendidas no farm** e as pessoas vendidas no **atacado** fornece somas desses números para cada linha e coluna.</span><span class="sxs-lookup"><span data-stu-id="47e8a-152">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="47e8a-153">No exemplo, **farm** e **tipo** são linhas, com as vendas de compra como os dados.</span><span class="sxs-lookup"><span data-stu-id="47e8a-153">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

![Uma tabela dinâmica mostrando as vendas totais de diferentes frutas com base no farm de onde elas vieram.](../images/excel-pivots-data-hierarchy.png)

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

## <a name="change-aggregation-function"></a><span data-ttu-id="47e8a-155">Função de agregação de alteração</span><span class="sxs-lookup"><span data-stu-id="47e8a-155">Change aggregation function</span></span>

<span data-ttu-id="47e8a-156">As hierarquias de dados têm seus valores agregados.</span><span class="sxs-lookup"><span data-stu-id="47e8a-156">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="47e8a-157">Para conjuntos de números de valores, esta é uma soma por padrão.</span><span class="sxs-lookup"><span data-stu-id="47e8a-157">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="47e8a-158">A `summarizeBy` propriedade define esse comportamento com base em um tipo [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) .</span><span class="sxs-lookup"><span data-stu-id="47e8a-158">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="47e8a-159">Os tipos de função de agregação `Sum`suportados `Count`atualmente `Average`são `Max`, `Min` `Product`,, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`,, `Automatic` e (o padrão).</span><span class="sxs-lookup"><span data-stu-id="47e8a-159">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="47e8a-160">O exemplo de código a seguir altera a agregação para ser a média dos dados.</span><span class="sxs-lookup"><span data-stu-id="47e8a-160">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="47e8a-161">Alterar cálculos com um ShowAsRule</span><span class="sxs-lookup"><span data-stu-id="47e8a-161">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="47e8a-162">As tabelas dinâmicas, por padrão, agregam os dados de suas hierarquias de linha e coluna de forma independente.</span><span class="sxs-lookup"><span data-stu-id="47e8a-162">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="47e8a-163">Um [ShowAsRule](/javascript/api/excel/excel.showasrule) altera a hierarquia de dados para valores de saída com base em outros itens na tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="47e8a-163">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="47e8a-164">O `ShowAsRule` objeto tem três propriedades:</span><span class="sxs-lookup"><span data-stu-id="47e8a-164">The `ShowAsRule` object has three properties:</span></span>

-   <span data-ttu-id="47e8a-165">`calculation`: O tipo de cálculo relativo a ser aplicado à hierarquia de dados (o padrão `none`é).</span><span class="sxs-lookup"><span data-stu-id="47e8a-165">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
-   <span data-ttu-id="47e8a-166">`baseField`: O campo dentro da hierarquia que contém os dados básicos antes do cálculo ser aplicado.</span><span class="sxs-lookup"><span data-stu-id="47e8a-166">`baseField`: The field within the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="47e8a-167">Normalmente [](/javascript/api/excel/excel.pivotfield) , o PivotField tem o mesmo nome de sua hierarquia pai.</span><span class="sxs-lookup"><span data-stu-id="47e8a-167">The [PivotField](/javascript/api/excel/excel.pivotfield) usually has the same name as its parent hierarchy.</span></span>
-   <span data-ttu-id="47e8a-168">`baseItem`: O [PivotItem](/javascript/api/excel/excel.pivotitem) individual comparado com os valores dos campos base com base no tipo de cálculo.</span><span class="sxs-lookup"><span data-stu-id="47e8a-168">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="47e8a-169">Nem todos os cálculos exigem esse campo.</span><span class="sxs-lookup"><span data-stu-id="47e8a-169">Not all calculations require this field.</span></span>

<span data-ttu-id="47e8a-170">O exemplo a seguir define o cálculo **da soma das** Enações vendidas na hierarquia de dados do farm como uma porcentagem do total da coluna.</span><span class="sxs-lookup"><span data-stu-id="47e8a-170">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span> <span data-ttu-id="47e8a-171">Ainda queremos que a granularidade seja estendida para o nível de tipo de frutas, portanto, usaremos a hierarquia de linha de **tipo** e seu campo base.</span><span class="sxs-lookup"><span data-stu-id="47e8a-171">We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field.</span></span> <span data-ttu-id="47e8a-172">O exemplo também tem o **farm** como a primeira hierarquia de linha, portanto, o total de entradas do farm exibe a porcentagem de produção de cada farm também.</span><span class="sxs-lookup"><span data-stu-id="47e8a-172">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![Uma tabela dinâmica mostrando as porcentagens das vendas de frutas em relação ao total geral de farms individuais e tipos de frutas individuais em cada farm.](../images/excel-pivots-showas-percentage.png)

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

<span data-ttu-id="47e8a-174">O exemplo anterior definiu o cálculo para a coluna, em relação a uma hierarquia de linha individual.</span><span class="sxs-lookup"><span data-stu-id="47e8a-174">The previous example set the calculation to the column, relative to an individual row hierarchy.</span></span> <span data-ttu-id="47e8a-175">Quando o cálculo está relacionado a um item individual, use a `baseItem` propriedade.</span><span class="sxs-lookup"><span data-stu-id="47e8a-175">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="47e8a-176">O exemplo a seguir mostra `differenceFrom` o cálculo.</span><span class="sxs-lookup"><span data-stu-id="47e8a-176">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="47e8a-177">Ele exibe a diferença entre as entradas de hierarquia de dados de vendas do farm em relação às de "farms".</span><span class="sxs-lookup"><span data-stu-id="47e8a-177">It displays the difference of the farm crate sales data hierarchy entries relative to those of “A Farms”.</span></span>
<span data-ttu-id="47e8a-178">O `baseField` **farm**de is, portanto, vemos as diferenças entre os outros farms, bem como as divisões de cada tipo de fruta (**Type** também é uma hierarquia de linha neste exemplo).</span><span class="sxs-lookup"><span data-stu-id="47e8a-178">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![Uma tabela dinâmica mostrando as diferenças das vendas de frutas entre "um farm" e outros.](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="pivottable-layouts"></a><span data-ttu-id="47e8a-182">Layouts de tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="47e8a-182">PivotTable layouts</span></span>

<span data-ttu-id="47e8a-183">Um [PivotLayout](/javascript/api/excel/excel.pivotlayout) define o posicionamento de hierarquias e seus dados.</span><span class="sxs-lookup"><span data-stu-id="47e8a-183">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="47e8a-184">Você acessa o layout para determinar os intervalos onde os dados são armazenados.</span><span class="sxs-lookup"><span data-stu-id="47e8a-184">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="47e8a-185">O diagrama a seguir mostra quais chamadas de função de layout correspondem aos intervalos da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="47e8a-185">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Um diagrama mostrando quais seções de uma tabela dinâmica são retornadas pelas funções obter intervalo do layout.](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="47e8a-187">O código a seguir demonstra como obter a última linha dos dados da tabela dinâmica percorrendo o layout.</span><span class="sxs-lookup"><span data-stu-id="47e8a-187">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="47e8a-188">Esses valores são somados em um total geral.</span><span class="sxs-lookup"><span data-stu-id="47e8a-188">Those values are then summed together for a grand total.</span></span>

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

<span data-ttu-id="47e8a-189">As tabelas dinâmicas têm três estilos de layout: compactar, estrutura de tópicos e tabular.</span><span class="sxs-lookup"><span data-stu-id="47e8a-189">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="47e8a-190">Vimos o estilo compacto nos exemplos anteriores.</span><span class="sxs-lookup"><span data-stu-id="47e8a-190">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="47e8a-191">Os exemplos a seguir usam os estilos de estrutura de tópicos e tabular, respectivamente.</span><span class="sxs-lookup"><span data-stu-id="47e8a-191">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="47e8a-192">O exemplo de código mostra como fazer o ciclo entre os diferentes layouts.</span><span class="sxs-lookup"><span data-stu-id="47e8a-192">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="47e8a-193">Layout de estrutura de tópicos</span><span class="sxs-lookup"><span data-stu-id="47e8a-193">Outline layout</span></span>

![Uma tabela dinâmica usando o layout de estrutura de tópicos.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="47e8a-195">Layout tabular</span><span class="sxs-lookup"><span data-stu-id="47e8a-195">Tabular layout</span></span>

![Uma tabela dinâmica usando o layout tabular.](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="47e8a-197">Alterar nomes de hierarquia</span><span class="sxs-lookup"><span data-stu-id="47e8a-197">Change hierarchy names</span></span>

<span data-ttu-id="47e8a-198">Os campos de hierarquia são editáveis.</span><span class="sxs-lookup"><span data-stu-id="47e8a-198">Hierarchy fields are editable.</span></span> <span data-ttu-id="47e8a-199">O código a seguir demonstra como alterar os nomes exibidos de duas hierarquias de dados.</span><span class="sxs-lookup"><span data-stu-id="47e8a-199">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="47e8a-200">Excluir uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="47e8a-200">Delete a PivotTable</span></span>

<span data-ttu-id="47e8a-201">As tabelas dinâmicas são excluídas usando seus nomes.</span><span class="sxs-lookup"><span data-stu-id="47e8a-201">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="47e8a-202">Confira também</span><span class="sxs-lookup"><span data-stu-id="47e8a-202">See also</span></span>

- [<span data-ttu-id="47e8a-203">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="47e8a-203">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="47e8a-204">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="47e8a-204">Excel JavaScript API Reference</span></span>](/javascript/api/excel)
