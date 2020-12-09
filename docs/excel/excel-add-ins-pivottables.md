---
title: Trabalhar com tabelas dinâmicas usando a API JavaScript do Excel
description: Use a API JavaScript do Excel para criar tabelas dinâmicas e interagir com seus componentes.
ms.date: 12/07/2020
localization_priority: Normal
ms.openlocfilehash: 0a1fefa6a855ab9ee1ccd71fd0dc60f282d2944b
ms.sourcegitcommit: fecad2afa7938d7178456c11ba52b558224813b4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/09/2020
ms.locfileid: "49603796"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="f70b0-103">Trabalhar com tabelas dinâmicas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f70b0-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="f70b0-104">As tabelas dinâmicas simplificam conjuntos de dados maiores.</span><span class="sxs-lookup"><span data-stu-id="f70b0-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="f70b0-105">Eles permitem a manipulação rápida dos dados agrupados.</span><span class="sxs-lookup"><span data-stu-id="f70b0-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="f70b0-106">A API JavaScript do Excel permite que o suplemento crie tabelas dinâmicas e interaja com seus componentes.</span><span class="sxs-lookup"><span data-stu-id="f70b0-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> <span data-ttu-id="f70b0-107">Este artigo descreve como as tabelas dinâmicas são representadas pela API JavaScript do Office e fornece exemplos de código para os principais cenários.</span><span class="sxs-lookup"><span data-stu-id="f70b0-107">This article describes how PivotTables are represented by the Office JavaScript API and provides code samples for key scenarios.</span></span>

<span data-ttu-id="f70b0-108">Se você não estiver familiarizado com a funcionalidade das tabelas dinâmicas, considere explorá-las como um usuário final.</span><span class="sxs-lookup"><span data-stu-id="f70b0-108">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span>
<span data-ttu-id="f70b0-109">Confira [criar uma tabela dinâmica para analisar os dados da planilha](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) para obter uma boa opção mais interessante nessas ferramentas.</span><span class="sxs-lookup"><span data-stu-id="f70b0-109">See [Create a PivotTable to analyze worksheet data](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f70b0-110">As tabelas dinâmicas criadas com OLAP não têm suporte no momento.</span><span class="sxs-lookup"><span data-stu-id="f70b0-110">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="f70b0-111">Também não há suporte para o Power pivot.</span><span class="sxs-lookup"><span data-stu-id="f70b0-111">There is also no support for Power Pivot.</span></span>

## <a name="object-model"></a><span data-ttu-id="f70b0-112">Modelo de objetos</span><span class="sxs-lookup"><span data-stu-id="f70b0-112">Object model</span></span>

<span data-ttu-id="f70b0-113">A [tabela dinâmica](/javascript/api/excel/excel.pivottable) é o objeto central para tabelas dinâmicas na API JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="f70b0-113">The [PivotTable](/javascript/api/excel/excel.pivottable) is the central object for PivotTables in the Office JavaScript API.</span></span>

- <span data-ttu-id="f70b0-114">`Workbook.pivotTables` e `Worksheet.pivotTables` são [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) que contêm as [tabelas dinâmicas](/javascript/api/excel/excel.pivottable) na pasta de trabalho e planilha, respectivamente.</span><span class="sxs-lookup"><span data-stu-id="f70b0-114">`Workbook.pivotTables` and `Worksheet.pivotTables` are [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that contain the [PivotTables](/javascript/api/excel/excel.pivottable) in the workbook and worksheet, respectively.</span></span>
- <span data-ttu-id="f70b0-115">Uma [tabela dinâmica](/javascript/api/excel/excel.pivottable) contém um [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) que tem vários [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span><span class="sxs-lookup"><span data-stu-id="f70b0-115">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) that has multiple [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span></span>
- <span data-ttu-id="f70b0-116">Esses [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) podem ser adicionados a coleções de hierarquias específicas para definir como os dados dinâmicos de tabela dinâmica (conforme explicado na [seção a seguir](#hierarchies)).</span><span class="sxs-lookup"><span data-stu-id="f70b0-116">These [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) can be added to specific hierarchy collections to define how the PivotTable pivots data (as explained in the [following section](#hierarchies)).</span></span>
- <span data-ttu-id="f70b0-117">Um [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contém um [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) que tem exatamente um [PivotField](/javascript/api/excel/excel.pivotfield).</span><span class="sxs-lookup"><span data-stu-id="f70b0-117">A [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contains a [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) that has exactly one [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="f70b0-118">Se o design expandir para incluir tabelas dinâmicas OLAP, isso pode ser alterado.</span><span class="sxs-lookup"><span data-stu-id="f70b0-118">If the design expands to include OLAP PivotTables, this may change.</span></span>
- <span data-ttu-id="f70b0-119">Um [PivotField](/javascript/api/excel/excel.pivotfield) pode ter um ou mais [PivotFilters](/javascript/api/excel/excel.pivotfilters) aplicados, desde que o [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) do campo seja atribuído a uma categoria de hierarquia.</span><span class="sxs-lookup"><span data-stu-id="f70b0-119">A [PivotField](/javascript/api/excel/excel.pivotfield) can have one or more [PivotFilters](/javascript/api/excel/excel.pivotfilters) applied, as long as the field's [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) is assigned to a hierarchy category.</span></span> 
- <span data-ttu-id="f70b0-120">Um [PivotField](/javascript/api/excel/excel.pivotfield) contém um [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) que tem vários [PivotItems](/javascript/api/excel/excel.pivotitem).</span><span class="sxs-lookup"><span data-stu-id="f70b0-120">A [PivotField](/javascript/api/excel/excel.pivotfield) contains a [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) that has multiple [PivotItems](/javascript/api/excel/excel.pivotitem).</span></span>
- <span data-ttu-id="f70b0-121">Uma [tabela dinâmica](/javascript/api/excel/excel.pivottable) contém um [PivotLayout](/javascript/api/excel/excel.pivotlayout) que define onde o [PivotFields](/javascript/api/excel/excel.pivotfield) e o [PivotItems](/javascript/api/excel/excel.pivotitem) são exibidos na planilha.</span><span class="sxs-lookup"><span data-stu-id="f70b0-121">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotLayout](/javascript/api/excel/excel.pivotlayout) that defines where the [PivotFields](/javascript/api/excel/excel.pivotfield) and [PivotItems](/javascript/api/excel/excel.pivotitem) are displayed in the worksheet.</span></span>

<span data-ttu-id="f70b0-122">Vamos ver como essas relações se aplicam a alguns dados de exemplo.</span><span class="sxs-lookup"><span data-stu-id="f70b0-122">Let's look at how these relationships apply to some example data.</span></span> <span data-ttu-id="f70b0-123">Os dados a seguir descrevem as vendas de frutas de vários farms.</span><span class="sxs-lookup"><span data-stu-id="f70b0-123">The following data describes fruit sales from various farms.</span></span> <span data-ttu-id="f70b0-124">Este será o exemplo neste artigo.</span><span class="sxs-lookup"><span data-stu-id="f70b0-124">It will be the example throughout this article.</span></span>

![Uma coleção de vendas de frutas de diferentes tipos de farms diferentes.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="f70b0-126">Estes dados de vendas do farm de frutas serão usados para criar uma tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="f70b0-126">This fruit farm sales data will be used to make a PivotTable.</span></span> <span data-ttu-id="f70b0-127">Cada coluna, como **tipos**, é um `PivotHierarchy` .</span><span class="sxs-lookup"><span data-stu-id="f70b0-127">Each column, such as **Types**, is a `PivotHierarchy`.</span></span> <span data-ttu-id="f70b0-128">A hierarquia **tipos** contém o campo **tipos** .</span><span class="sxs-lookup"><span data-stu-id="f70b0-128">The **Types** hierarchy contains the **Types** field.</span></span> <span data-ttu-id="f70b0-129">O campo **tipos** contém os itens **Apple**, **Kiwi**, **casca**, **verde-limão** e **laranja**.</span><span class="sxs-lookup"><span data-stu-id="f70b0-129">The **Types** field contains the items **Apple**, **Kiwi**, **Lemon**, **Lime**, and **Orange**.</span></span>

### <a name="hierarchies"></a><span data-ttu-id="f70b0-130">Hierarquias</span><span class="sxs-lookup"><span data-stu-id="f70b0-130">Hierarchies</span></span>

<span data-ttu-id="f70b0-131">As tabelas dinâmicas são organizadas com base em quatro categorias de hierarquia: [linha](/javascript/api/excel/excel.rowcolumnpivothierarchy), [coluna](/javascript/api/excel/excel.rowcolumnpivothierarchy), [dados](/javascript/api/excel/excel.datapivothierarchy)e [filtro](/javascript/api/excel/excel.filterpivothierarchy).</span><span class="sxs-lookup"><span data-stu-id="f70b0-131">PivotTables are organized based on four hierarchy categories: [row](/javascript/api/excel/excel.rowcolumnpivothierarchy), [column](/javascript/api/excel/excel.rowcolumnpivothierarchy), [data](/javascript/api/excel/excel.datapivothierarchy), and [filter](/javascript/api/excel/excel.filterpivothierarchy).</span></span>

<span data-ttu-id="f70b0-132">Os dados do farm mostrados anteriormente têm cinco hierarquias: **farms**, **tipo**, **classificação**, enessações **vendidas no farm** e as **dovendas vendidas no atacado**.</span><span class="sxs-lookup"><span data-stu-id="f70b0-132">The farm data shown earlier has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="f70b0-133">Cada hierarquia só pode existir em uma das quatro categorias.</span><span class="sxs-lookup"><span data-stu-id="f70b0-133">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="f70b0-134">Se **Type** for adicionado às hierarquias de coluna, ele também não poderá estar na linha, dados ou hierarquias de filtro.</span><span class="sxs-lookup"><span data-stu-id="f70b0-134">If **Type** is added to column hierarchies, it cannot also be in the row, data, or filter hierarchies.</span></span> <span data-ttu-id="f70b0-135">Se **Type** for adicionado posteriormente às hierarquias de linha, ele será removido das hierarquias de coluna.</span><span class="sxs-lookup"><span data-stu-id="f70b0-135">If **Type** is subsequently added to row hierarchies, it is removed from the column hierarchies.</span></span> <span data-ttu-id="f70b0-136">Esse comportamento é o mesmo que a atribuição de hierarquia é feita por meio da interface do usuário do Excel ou das APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="f70b0-136">This behavior is the same whether hierarchy assignment is done through the Excel UI or the Excel JavaScript APIs.</span></span>

<span data-ttu-id="f70b0-137">Hierarquias de linha e coluna definem como os dados serão agrupados.</span><span class="sxs-lookup"><span data-stu-id="f70b0-137">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="f70b0-138">Por exemplo, uma hierarquia de linha de **farms** agrupará todos os conjuntos de dados do mesmo farm.</span><span class="sxs-lookup"><span data-stu-id="f70b0-138">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="f70b0-139">A escolha entre hierarquia de linha e coluna define a orientação da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="f70b0-139">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="f70b0-140">Hierarquias de dados são os valores a serem agregados com base nas hierarquias de linha e coluna.</span><span class="sxs-lookup"><span data-stu-id="f70b0-140">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="f70b0-141">Uma tabela dinâmica com uma hierarquia de linha de **farms** e uma hierarquia de dados de **envenda vendida** mostra a soma total (por padrão) de todos os diferentes frutas para cada farm.</span><span class="sxs-lookup"><span data-stu-id="f70b0-141">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="f70b0-142">As hierarquias de filtro incluem ou excluem dados da tabela dinâmica com base nos valores desse tipo filtrado.</span><span class="sxs-lookup"><span data-stu-id="f70b0-142">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="f70b0-143">Uma hierarquia de filtro de **classificação** com o tipo **orgânica** selecionado mostra apenas dados para frutas orgânicas.</span><span class="sxs-lookup"><span data-stu-id="f70b0-143">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="f70b0-144">Estes são os dados do farm novamente, juntamente com uma tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="f70b0-144">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="f70b0-145">A tabela dinâmica está usando o **farm** e o **tipo** como hierarquias de linha, as enfileiras **vendidas no farm** e as doutilizações **vendidas** como as hierarquias de dados (com a função de agregação padrão de Sum) e a **classificação** como uma hierarquia de filtro (com a **orgânica** selecionada).</span><span class="sxs-lookup"><span data-stu-id="f70b0-145">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span>

![Uma seleção de dados de vendas de frutas ao lado de uma tabela dinâmica com hierarquias de linha, dados e filtros.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="f70b0-147">Esta tabela dinâmica pode ser gerada por meio da API JavaScript ou através da interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="f70b0-147">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="f70b0-148">Ambas as opções permitem mais manipulação por meio de suplementos.</span><span class="sxs-lookup"><span data-stu-id="f70b0-148">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="f70b0-149">Criar uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="f70b0-149">Create a PivotTable</span></span>

<span data-ttu-id="f70b0-150">As tabelas dinâmicas precisam de um nome, origem e destino.</span><span class="sxs-lookup"><span data-stu-id="f70b0-150">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="f70b0-151">A origem pode ser um endereço de intervalo ou nome de tabela (passado como um `Range` , `string` ou `Table` tipo).</span><span class="sxs-lookup"><span data-stu-id="f70b0-151">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="f70b0-152">O destino é um endereço de intervalo (fornecido como `Range` ou um ou `string` ).</span><span class="sxs-lookup"><span data-stu-id="f70b0-152">The destination is a range address (given as either a `Range` or `string`).</span></span>
<span data-ttu-id="f70b0-153">Os exemplos a seguir mostram várias técnicas de criação de tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="f70b0-153">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="f70b0-154">Criar uma tabela dinâmica com endereços de intervalo</span><span class="sxs-lookup"><span data-stu-id="f70b0-154">Create a PivotTable with range addresses</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="f70b0-155">Criar uma tabela dinâmica com objetos Range</span><span class="sxs-lookup"><span data-stu-id="f70b0-155">Create a PivotTable with Range objects</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21.
    var rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    var rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add(
      "Farm Sales", rangeToAnalyze, rangeToPlacePivot);

    return context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="f70b0-156">Criar uma tabela dinâmica no nível da pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="f70b0-156">Create a PivotTable at the workbook level</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="f70b0-157">Usar uma tabela dinâmica existente</span><span class="sxs-lookup"><span data-stu-id="f70b0-157">Use an existing PivotTable</span></span>

<span data-ttu-id="f70b0-158">As tabelas dinâmicas criadas manualmente também podem ser acessadas por meio da coleção PivotTable da pasta de trabalho ou de planilhas individuais.</span><span class="sxs-lookup"><span data-stu-id="f70b0-158">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> <span data-ttu-id="f70b0-159">O código a seguir obtém uma tabela dinâmica chamada **My pivot** da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="f70b0-159">The following code gets a PivotTable named **My Pivot** from the workbook.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="f70b0-160">Adicionar linhas e colunas a uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="f70b0-160">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="f70b0-161">Linhas e colunas dinamizam os dados em torno dos valores dos campos.</span><span class="sxs-lookup"><span data-stu-id="f70b0-161">Rows and columns pivot the data around those fields' values.</span></span>

<span data-ttu-id="f70b0-162">A adição da coluna do **farm** dinamiza todas as vendas em torno de cada farm.</span><span class="sxs-lookup"><span data-stu-id="f70b0-162">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="f70b0-163">Adicionar as linhas de **tipo** e **classificação** divide ainda mais os dados com base no que frutas foi vendido e se foi orgânica ou não.</span><span class="sxs-lookup"><span data-stu-id="f70b0-163">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

![Uma tabela dinâmica com uma coluna do farm e linhas de tipo e classificação.](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

<span data-ttu-id="f70b0-165">Você também pode ter uma tabela dinâmica com apenas linhas ou colunas.</span><span class="sxs-lookup"><span data-stu-id="f70b0-165">You can also have a PivotTable with only rows or columns.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="f70b0-166">Adicionar hierarquias de dados à tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="f70b0-166">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="f70b0-167">As hierarquias de dados preenchem a tabela dinâmica com informações para combinar com base nas linhas e colunas.</span><span class="sxs-lookup"><span data-stu-id="f70b0-167">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="f70b0-168">Adicionar as hierarquias de dados das pessoas **vendidas no farm** e as pessoas **vendidas no atacado** fornece somas desses números para cada linha e coluna.</span><span class="sxs-lookup"><span data-stu-id="f70b0-168">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span>

<span data-ttu-id="f70b0-169">No exemplo, **farm** e **tipo** são linhas, com as vendas de compra como os dados.</span><span class="sxs-lookup"><span data-stu-id="f70b0-169">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span>

![Uma tabela dinâmica mostrando as vendas totais de diferentes frutas com base no farm de onde elas vieram.](../images/excel-pivots-data-hierarchy.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based.
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the hierarchies
    // that will have their data aggregated (summed in this case).
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    return context.sync();
});
```

## <a name="pivottable-layouts-and-getting-pivoted-data"></a><span data-ttu-id="f70b0-171">Layouts de tabela dinâmica e obtendo dados dinâmicos</span><span class="sxs-lookup"><span data-stu-id="f70b0-171">PivotTable layouts and getting pivoted data</span></span>

<span data-ttu-id="f70b0-172">Um [PivotLayout](/javascript/api/excel/excel.pivotlayout) define o posicionamento de hierarquias e seus dados.</span><span class="sxs-lookup"><span data-stu-id="f70b0-172">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="f70b0-173">Você acessa o layout para determinar os intervalos onde os dados são armazenados.</span><span class="sxs-lookup"><span data-stu-id="f70b0-173">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="f70b0-174">O diagrama a seguir mostra quais chamadas de função de layout correspondem aos intervalos da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="f70b0-174">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Um diagrama mostrando quais seções de uma tabela dinâmica são retornadas pelas funções obter intervalo do layout.](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a><span data-ttu-id="f70b0-176">Obter dados da tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="f70b0-176">Get data from the PivotTable</span></span>

<span data-ttu-id="f70b0-177">O layout define como a tabela dinâmica é exibida na planilha.</span><span class="sxs-lookup"><span data-stu-id="f70b0-177">The layout defines how the PivotTable is displayed in the worksheet.</span></span> <span data-ttu-id="f70b0-178">Isso significa que o `PivotLayout` objeto controla os intervalos usados para elementos de tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="f70b0-178">This means the `PivotLayout` object controls the ranges used for PivotTable elements.</span></span> <span data-ttu-id="f70b0-179">Use os intervalos fornecidos pelo layout para obter dados coletados e agregados pela tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="f70b0-179">Use the ranges provided by the layout to get data collected and aggregated by the PivotTable.</span></span> <span data-ttu-id="f70b0-180">Em particular, use `PivotLayout.getDataBodyRange` para acessar o que a tabela dinâmica produz.</span><span class="sxs-lookup"><span data-stu-id="f70b0-180">In particular, use `PivotLayout.getDataBodyRange` to access what the PivotTable produces.</span></span>

<span data-ttu-id="f70b0-181">O código a seguir demonstra como obter a última linha dos dados da tabela dinâmica percorrendo o layout (o **total geral** da soma de enfileiras **vendidas no farm** e a **soma das colunas vendidas do atacadista** no exemplo anterior).</span><span class="sxs-lookup"><span data-stu-id="f70b0-181">The following code demonstrates how to get the last row of the PivotTable data by going through the layout (the **Grand Total** of both the **Sum of Crates Sold at Farm** and **Sum of Crates Sold Wholesale** columns in the earlier example).</span></span> <span data-ttu-id="f70b0-182">Esses valores são somados em um total final, que é exibido na célula **E30** (fora da tabela dinâmica).</span><span class="sxs-lookup"><span data-stu-id="f70b0-182">Those values are then summed together for a final total, which is displayed in cell **E30** (outside of the PivotTable).</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // Get the totals for each data hierarchy from the layout.
    var range = pivotTable.layout.getDataBodyRange();
    var grandTotalRange = range.getLastRow();
    grandTotalRange.load("address");
    return context.sync().then(function () {
        // Sum the totals from the PivotTable data hierarchies and place them in a new range, outside of the PivotTable.
        var masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("E30");
        masterTotalRange.formulas = [["=SUM(" + grandTotalRange.address + ")"]];
    });
});
```

### <a name="layout-types"></a><span data-ttu-id="f70b0-183">Tipos de layout</span><span class="sxs-lookup"><span data-stu-id="f70b0-183">Layout types</span></span>

<span data-ttu-id="f70b0-184">As tabelas dinâmicas têm três estilos de layout: compactar, estrutura de tópicos e tabular.</span><span class="sxs-lookup"><span data-stu-id="f70b0-184">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="f70b0-185">Vimos o estilo compacto nos exemplos anteriores.</span><span class="sxs-lookup"><span data-stu-id="f70b0-185">We've seen the compact style in the previous examples.</span></span>

<span data-ttu-id="f70b0-186">Os exemplos a seguir usam os estilos de estrutura de tópicos e tabular, respectivamente.</span><span class="sxs-lookup"><span data-stu-id="f70b0-186">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="f70b0-187">O exemplo de código mostra como fazer o ciclo entre os diferentes layouts.</span><span class="sxs-lookup"><span data-stu-id="f70b0-187">The code sample shows how to cycle between the different layouts.</span></span>

#### <a name="outline-layout"></a><span data-ttu-id="f70b0-188">Layout de estrutura de tópicos</span><span class="sxs-lookup"><span data-stu-id="f70b0-188">Outline layout</span></span>

![Uma tabela dinâmica usando o layout de estrutura de tópicos.](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a><span data-ttu-id="f70b0-190">Layout tabular</span><span class="sxs-lookup"><span data-stu-id="f70b0-190">Tabular layout</span></span>

![Uma tabela dinâmica usando o layout tabular.](../images/excel-pivots-tabular-layout.png)

## <a name="delete-a-pivottable"></a><span data-ttu-id="f70b0-192">Excluir uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="f70b0-192">Delete a PivotTable</span></span>

<span data-ttu-id="f70b0-193">As tabelas dinâmicas são excluídas usando seus nomes.</span><span class="sxs-lookup"><span data-stu-id="f70b0-193">PivotTables are deleted by using their name.</span></span>

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a><span data-ttu-id="f70b0-194">Filtrar uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="f70b0-194">Filter a PivotTable</span></span>

<span data-ttu-id="f70b0-195">O método principal para filtrar dados de tabela dinâmica é com o PivotFilters.</span><span class="sxs-lookup"><span data-stu-id="f70b0-195">The primary method for filtering PivotTable data is with PivotFilters.</span></span> <span data-ttu-id="f70b0-196">Slicers oferecem um método de filtragem alternativo e menos flexível.</span><span class="sxs-lookup"><span data-stu-id="f70b0-196">Slicers offer an alternate, less flexible filtering method.</span></span> 

<span data-ttu-id="f70b0-197">[PivotFilters](/javascript/api/excel/excel.pivotfilters) filtrar dados com base em quatro categorias de [hierarquia](#hierarchies) de tabela dinâmica (filtros, colunas, linhas e valores).</span><span class="sxs-lookup"><span data-stu-id="f70b0-197">[PivotFilters](/javascript/api/excel/excel.pivotfilters) filter data based on a PivotTable's four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="f70b0-198">Há quatro tipos de PivotFilters, permitindo filtragem baseada em data de calendário, análise de cadeia de caracteres, comparação de números e filtragem com base em uma entrada personalizada.</span><span class="sxs-lookup"><span data-stu-id="f70b0-198">There are four types of PivotFilters, allowing calendar date-based filtering, string parsing, number comparison, and filtering based on a custom input.</span></span> 

<span data-ttu-id="f70b0-199">[Slicers](/javascript/api/excel/excel.slicer) podem ser aplicados a tabelas dinâmicas e tabelas regulares do Excel.</span><span class="sxs-lookup"><span data-stu-id="f70b0-199">[Slicers](/javascript/api/excel/excel.slicer) can be applied to both PivotTables and regular Excel tables.</span></span> <span data-ttu-id="f70b0-200">Quando aplicado a uma tabela dinâmica, a segmentação de dados funciona como um [PivotManualFilter](#pivotmanualfilter) e permite a filtragem com base em uma entrada personalizada.</span><span class="sxs-lookup"><span data-stu-id="f70b0-200">When applied to a PivotTable, slicers function like a [PivotManualFilter](#pivotmanualfilter) and allow filtering based on a custom input.</span></span> <span data-ttu-id="f70b0-201">Ao contrário do PivotFilters, as segmentações de conteúdo têm um [componente de UI do Excel](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d).</span><span class="sxs-lookup"><span data-stu-id="f70b0-201">Unlike PivotFilters, slicers have an [Excel UI component](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d).</span></span> <span data-ttu-id="f70b0-202">Com a `Slicer` turma, você cria esse componente de interface do usuário, gerencia a filtragem e controla sua aparência visual.</span><span class="sxs-lookup"><span data-stu-id="f70b0-202">With the `Slicer` class, you create this UI component, manage filtering, and control its visual appearance.</span></span> 

### <a name="filter-with-pivotfilters"></a><span data-ttu-id="f70b0-203">Filtro com PivotFilters</span><span class="sxs-lookup"><span data-stu-id="f70b0-203">Filter with PivotFilters</span></span>

<span data-ttu-id="f70b0-204">O [PivotFilters](/javascript/api/excel/excel.pivotfilters) permite que você filtre os dados da tabela dinâmica com base nas quatro [categorias de hierarquia](#hierarchies) (filtros, colunas, linhas e valores).</span><span class="sxs-lookup"><span data-stu-id="f70b0-204">[PivotFilters](/javascript/api/excel/excel.pivotfilters) allow you to filter PivotTable data based on the four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="f70b0-205">No modelo de objeto de tabela dinâmica, `PivotFilters` são aplicadas a um [PivotField](/javascript/api/excel/excel.pivotfield), e cada `PivotField` uma pode ter uma ou mais atribuições `PivotFilters` .</span><span class="sxs-lookup"><span data-stu-id="f70b0-205">In the PivotTable object model, `PivotFilters` are applied to a [PivotField](/javascript/api/excel/excel.pivotfield), and each `PivotField` can have one or more assigned `PivotFilters`.</span></span> <span data-ttu-id="f70b0-206">Para aplicar PivotFilters a um PivotField, o [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) correspondente do campo deve ser atribuído a uma categoria de hierarquia.</span><span class="sxs-lookup"><span data-stu-id="f70b0-206">To apply PivotFilters to a PivotField, the field's corresponding [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) must be assigned to a hierarchy category.</span></span> 

#### <a name="types-of-pivotfilters"></a><span data-ttu-id="f70b0-207">Tipos de PivotFilters</span><span class="sxs-lookup"><span data-stu-id="f70b0-207">Types of PivotFilters</span></span>

| <span data-ttu-id="f70b0-208">Tipo de filtro</span><span class="sxs-lookup"><span data-stu-id="f70b0-208">Filter type</span></span> | <span data-ttu-id="f70b0-209">Finalidade do filtro</span><span class="sxs-lookup"><span data-stu-id="f70b0-209">Filter purpose</span></span> | <span data-ttu-id="f70b0-210">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f70b0-210">Excel JavaScript API reference</span></span> |
|:--- |:--- |:--- |
| <span data-ttu-id="f70b0-211">DateFilter</span><span class="sxs-lookup"><span data-stu-id="f70b0-211">DateFilter</span></span> | <span data-ttu-id="f70b0-212">Filtragem baseada em data do calendário.</span><span class="sxs-lookup"><span data-stu-id="f70b0-212">Calendar date-based filtering.</span></span> | [<span data-ttu-id="f70b0-213">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="f70b0-213">PivotDateFilter</span></span>](/javascript/api/excel/excel.pivotdatefilter) |
| <span data-ttu-id="f70b0-214">LabelFilter</span><span class="sxs-lookup"><span data-stu-id="f70b0-214">LabelFilter</span></span> | <span data-ttu-id="f70b0-215">Filtragem de comparação de texto.</span><span class="sxs-lookup"><span data-stu-id="f70b0-215">Text comparison filtering.</span></span> | [<span data-ttu-id="f70b0-216">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="f70b0-216">PivotLabelFilter</span></span>](/javascript/api/excel/excel.pivotlabelfilter) |
| <span data-ttu-id="f70b0-217">ManualFilter</span><span class="sxs-lookup"><span data-stu-id="f70b0-217">ManualFilter</span></span> | <span data-ttu-id="f70b0-218">Filtragem de entrada personalizada.</span><span class="sxs-lookup"><span data-stu-id="f70b0-218">Custom input filtering.</span></span> | [<span data-ttu-id="f70b0-219">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="f70b0-219">PivotManualFilter</span></span>](/javascript/api/excel/excel.pivotmanualfilter) |
| <span data-ttu-id="f70b0-220">ValueFilter</span><span class="sxs-lookup"><span data-stu-id="f70b0-220">ValueFilter</span></span> | <span data-ttu-id="f70b0-221">Filtragem de comparação de número.</span><span class="sxs-lookup"><span data-stu-id="f70b0-221">Number comparison filtering.</span></span> | [<span data-ttu-id="f70b0-222">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="f70b0-222">PivotValueFilter</span></span>](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a><span data-ttu-id="f70b0-223">Criar um PivotFilter</span><span class="sxs-lookup"><span data-stu-id="f70b0-223">Create a PivotFilter</span></span>

<span data-ttu-id="f70b0-224">Para filtrar os dados da tabela dinâmica com um filtro dinâmico \* (como um PivotDateFilter), aplique o filtro a um [PivotField](/javascript/api/excel/excel.pivotfield).</span><span class="sxs-lookup"><span data-stu-id="f70b0-224">To filter PivotTable data with a Pivot\*Filter (such as a PivotDateFilter), apply the filter to a [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="f70b0-225">Os quatro exemplos de código a seguir mostram como usar cada um dos quatro tipos de PivotFilters.</span><span class="sxs-lookup"><span data-stu-id="f70b0-225">The following four code samples show how to use each of the four types of PivotFilters.</span></span> 

##### <a name="pivotdatefilter"></a><span data-ttu-id="f70b0-226">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="f70b0-226">PivotDateFilter</span></span>

<span data-ttu-id="f70b0-227">O primeiro exemplo de código aplica um [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) à **Data updated** PivotField, ocultando quaisquer dados antes de **2020-08-01**.</span><span class="sxs-lookup"><span data-stu-id="f70b0-227">The first code sample applies a [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) to the **Date Updated** PivotField, hiding any data prior to **2020-08-01**.</span></span> 

> [!IMPORTANT] 
> <span data-ttu-id="f70b0-228">Um filtro pivot \* não pode ser aplicado a um PivotField, a menos que o PivotHierarchy do campo seja atribuído a uma categoria de hierarquia.</span><span class="sxs-lookup"><span data-stu-id="f70b0-228">A Pivot\*Filter can't be applied to a PivotField unless that field's PivotHierarchy is assigned to a hierarchy category.</span></span> <span data-ttu-id="f70b0-229">No exemplo de código a seguir, o `dateHierarchy` deve ser adicionado à categoria da tabela dinâmica para `rowHierarchies` que possa ser usado para filtragem.</span><span class="sxs-lookup"><span data-stu-id="f70b0-229">In the following code sample, the `dateHierarchy` must be added to the PivotTable's `rowHierarchies` category before it can be used for filtering.</span></span>

```js
Excel.run(function (context) {
    // Get the PivotTable and the date hierarchy.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var dateHierarchy = pivotTable.rowHierarchies.getItemOrNullObject("Date Updated");
    
    return context.sync().then(function () {
        // PivotFilters can only be applied to PivotHierarchies that are being used for pivoting.
        // If it's not already there, add "Date Updated" to the hierarchies.
        if (dateHierarchy.isNullObject) {
          dateHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Date Updated"));
        }

        // Apply a date filter to filter out anything logged before August.
        var filterField = dateHierarchy.fields.getItem("Date Updated");
        var dateFilter = {
          condition: Excel.DateFilterCondition.afterOrEqualTo,
          comparator: {
            date: "2020-08-01",
            specificity: Excel.FilterDatetimeSpecificity.month
          }
        };
        filterField.applyFilter({ dateFilter: dateFilter });
        
        return context.sync();
    });
});
```

> [!NOTE]
> <span data-ttu-id="f70b0-230">Os três trechos de código a seguir exibem apenas trechos de filtro específico, em vez de `Excel.run` chamadas completas.</span><span class="sxs-lookup"><span data-stu-id="f70b0-230">The following three code snippets only display filter-specific excerpts, instead of full `Excel.run` calls.</span></span>

##### <a name="pivotlabelfilter"></a><span data-ttu-id="f70b0-231">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="f70b0-231">PivotLabelFilter</span></span>

<span data-ttu-id="f70b0-232">O segundo trecho de código demonstra como aplicar [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) ao **tipo** PivotField, usando a `LabelFilterCondition.beginsWith` propriedade para excluir rótulos que começam com a letra **L**.</span><span class="sxs-lookup"><span data-stu-id="f70b0-232">The second code snippet demonstrates how to apply a [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) to the **Type** PivotField, using the `LabelFilterCondition.beginsWith` property to exclude labels that start with the letter **L**.</span></span> 

```js
    // Get the "Type" field.
    var filterField = pivotTable.hierarchies.getItem("Type").fields.getItem("Type");

    // Filter out any types that start with "L" ("Lemons" and "Limes" in this case).
    var filter: Excel.PivotLabelFilter = {
      condition: Excel.LabelFilterCondition.beginsWith,
      substring: "L",
      exclusive: true
    };

    // Apply the label filter to the field.
    filterField.applyFilter({ labelFilter: filter });
```

##### <a name="pivotmanualfilter"></a><span data-ttu-id="f70b0-233">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="f70b0-233">PivotManualFilter</span></span>

<span data-ttu-id="f70b0-234">O terceiro trecho de código aplica um filtro manual com [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) ao campo de **classificação** , filtrando dados que não incluem a classificação **orgânica**.</span><span class="sxs-lookup"><span data-stu-id="f70b0-234">The third code snippet applies a manual filter with [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) to the the **Classification** field, filtering out data that doesn't include the classification **Organic**.</span></span> 

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a><span data-ttu-id="f70b0-235">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="f70b0-235">PivotValueFilter</span></span>

<span data-ttu-id="f70b0-236">Para comparar números, use um filtro de valor com [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter), conforme mostrado no trecho de código final.</span><span class="sxs-lookup"><span data-stu-id="f70b0-236">To compare numbers, use a value filter with [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter), as shown in the final code snippet.</span></span> <span data-ttu-id="f70b0-237">O `PivotValueFilter` compara os dados no PivotField do **farm** com os dados no PivotField **atacadista vendido** , incluindo apenas farms cuja soma das enessações vendidas excede o valor **500**.</span><span class="sxs-lookup"><span data-stu-id="f70b0-237">The `PivotValueFilter` compares the data in the **Farm** PivotField to the data in the **Crates Sold Wholesale** PivotField, including only farms whose sum of crates sold exceeds the value **500**.</span></span> 

```js
    // Get the "Farm" field.
    var filterField = pivotTable.hierarchies.getItem("Farm").fields.getItem("Farm");
    
    // Filter to only include rows with more than 500 wholesale crates sold.
    var filter: Excel.PivotValueFilter = {
      condition: Excel.ValueFilterCondition.greaterThan,
      comparator: 500,
      value: "Sum of Crates Sold Wholesale"
    };
    
    // Apply the value filter to the field.
    filterField.applyFilter({ valueFilter: filter });
```

#### <a name="remove-pivotfilters"></a><span data-ttu-id="f70b0-238">Remover PivotFilters</span><span class="sxs-lookup"><span data-stu-id="f70b0-238">Remove PivotFilters</span></span>

<span data-ttu-id="f70b0-239">Para remover todos os PivotFilters, aplique o `clearAllFilters` método a cada PivotField, conforme mostrado no exemplo de código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f70b0-239">To remove all PivotFilters, apply the `clearAllFilters` method to each PivotField, as shown in the following code sample.</span></span> 

```js
Excel.run(function (context) {
    // Get the PivotTable.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.hierarchies.load("name");
    
    return context.sync().then(function () {
        // Clear the filters on each PivotField.
        pivotTable.hierarchies.items.forEach(function (hierarchy) {
          hierarchy.fields.getItem(hierarchy.name).clearAllFilters();
        });
        return context.sync();
    });
});
```

### <a name="filter-with-slicers"></a><span data-ttu-id="f70b0-240">Filtro com segmentação de Slice</span><span class="sxs-lookup"><span data-stu-id="f70b0-240">Filter with slicers</span></span>

<span data-ttu-id="f70b0-241">As [segmentações](/javascript/api/excel/excel.slicer) de dados permitem que os dados sejam filtrados de uma tabela dinâmica ou tabela do Excel.</span><span class="sxs-lookup"><span data-stu-id="f70b0-241">[Slicers](/javascript/api/excel/excel.slicer) allow data to be filtered from an Excel PivotTable or table.</span></span> <span data-ttu-id="f70b0-242">Uma segmentação de, usa valores de uma coluna especificada ou PivotField para filtrar as linhas correspondentes.</span><span class="sxs-lookup"><span data-stu-id="f70b0-242">A slicer uses values from a specified column or PivotField to filter corresponding rows.</span></span> <span data-ttu-id="f70b0-243">Esses valores são armazenados como objetos [SlicerItem](/javascript/api/excel/excel.sliceritem) no `Slicer` .</span><span class="sxs-lookup"><span data-stu-id="f70b0-243">These values are stored as [SlicerItem](/javascript/api/excel/excel.sliceritem) objects in the `Slicer`.</span></span> <span data-ttu-id="f70b0-244">O suplemento pode ajustar esses filtros, como os usuários ([por meio da interface do usuário do Excel](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span><span class="sxs-lookup"><span data-stu-id="f70b0-244">Your add-in can adjust these filters, as can users ([through the Excel UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span></span> <span data-ttu-id="f70b0-245">A segmentação de trabalho fica na parte superior da planilha na camada de desenho, conforme mostrado na captura de tela a seguir.</span><span class="sxs-lookup"><span data-stu-id="f70b0-245">The slicer sits on top of the worksheet in the drawing layer, as shown in the following screenshot.</span></span>

![Uma segmentação de dados Filtrando dados em uma tabela dinâmica.](../images/excel-slicer.png)

> [!NOTE]
> <span data-ttu-id="f70b0-247">As técnicas descritas nesta seção concentram-se em como usar slicers conectados a tabelas dinâmicas.</span><span class="sxs-lookup"><span data-stu-id="f70b0-247">The techniques described in this section focus on how to use slicers connected to PivotTables.</span></span> <span data-ttu-id="f70b0-248">As mesmas técnicas também se aplicam ao uso de segmentações de, conectadas a tabelas.</span><span class="sxs-lookup"><span data-stu-id="f70b0-248">The same techniques also apply to using slicers connected to tables.</span></span>

#### <a name="create-a-slicer"></a><span data-ttu-id="f70b0-249">Criar uma segmentação de um</span><span class="sxs-lookup"><span data-stu-id="f70b0-249">Create a slicer</span></span>

<span data-ttu-id="f70b0-250">Você pode criar uma segmentação de, em uma pasta de trabalho ou planilha, usando o `Workbook.slicers.add` método ou `Worksheet.slicers.add` método.</span><span class="sxs-lookup"><span data-stu-id="f70b0-250">You can create a slicer in a workbook or worksheet by using the `Workbook.slicers.add` method or `Worksheet.slicers.add` method.</span></span> <span data-ttu-id="f70b0-251">Isso adiciona uma segmentação de objetos à [SlicerCollection](/javascript/api/excel/excel.slicercollection) do objeto especificado `Workbook` ou `Worksheet` .</span><span class="sxs-lookup"><span data-stu-id="f70b0-251">Doing so adds a slicer to the [SlicerCollection](/javascript/api/excel/excel.slicercollection) of the specified `Workbook` or `Worksheet` object.</span></span> <span data-ttu-id="f70b0-252">O `SlicerCollection.add` método tem três parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f70b0-252">The `SlicerCollection.add` method has three parameters:</span></span>

- <span data-ttu-id="f70b0-253">`slicerSource`: A fonte de dados na qual a nova segmentação de dados se baseia.</span><span class="sxs-lookup"><span data-stu-id="f70b0-253">`slicerSource`: The data source on which the new slicer is based.</span></span> <span data-ttu-id="f70b0-254">Pode ser um `PivotTable` , `Table` ou cadeia de caracteres que representa o nome ou a ID de um `PivotTable` ou `Table` .</span><span class="sxs-lookup"><span data-stu-id="f70b0-254">It can be a `PivotTable`, `Table`, or string representing the name or ID of a `PivotTable` or `Table`.</span></span>
- <span data-ttu-id="f70b0-255">`sourceField`: O campo na fonte de dados pela qual filtrar.</span><span class="sxs-lookup"><span data-stu-id="f70b0-255">`sourceField`: The field in the data source by which to filter.</span></span> <span data-ttu-id="f70b0-256">Pode ser um `PivotField` , `TableColumn` ou cadeia de caracteres que representa o nome ou a ID de um `PivotField` ou `TableColumn` .</span><span class="sxs-lookup"><span data-stu-id="f70b0-256">It can be a `PivotField`, `TableColumn`, or string representing the name or ID of a `PivotField` or `TableColumn`.</span></span>
- <span data-ttu-id="f70b0-257">`slicerDestination`: A planilha onde a nova segmentação de trabalho será criada.</span><span class="sxs-lookup"><span data-stu-id="f70b0-257">`slicerDestination`: The worksheet where the new slicer will be created.</span></span> <span data-ttu-id="f70b0-258">Pode ser um `Worksheet` objeto ou o nome ou a ID de um `Worksheet` .</span><span class="sxs-lookup"><span data-stu-id="f70b0-258">It can be a `Worksheet` object or the name or ID of a `Worksheet`.</span></span> <span data-ttu-id="f70b0-259">Esse parâmetro é desnecessário quando o `SlicerCollection` é acessado `Worksheet.slicers` .</span><span class="sxs-lookup"><span data-stu-id="f70b0-259">This parameter is unnecessary when the `SlicerCollection` is accessed through `Worksheet.slicers`.</span></span> <span data-ttu-id="f70b0-260">Nesse caso, a planilha da coleção é usada como o destino.</span><span class="sxs-lookup"><span data-stu-id="f70b0-260">In this case, the collection's worksheet is used as the destination.</span></span>

<span data-ttu-id="f70b0-261">O exemplo de código a seguir adiciona uma nova segmentação de trabalho à planilha **dinâmica** .</span><span class="sxs-lookup"><span data-stu-id="f70b0-261">The following code sample adds a new slicer to the **Pivot** worksheet.</span></span> <span data-ttu-id="f70b0-262">A origem da segmentação de dados é a tabela dinâmica de **vendas do farm** e filtra usando os dados do **tipo** .</span><span class="sxs-lookup"><span data-stu-id="f70b0-262">The slicer's source is the **Farm Sales** PivotTable and filters using the **Type** data.</span></span> <span data-ttu-id="f70b0-263">A segmentação de, também é chamada de **segmentação de frutas** para referência futura.</span><span class="sxs-lookup"><span data-stu-id="f70b0-263">The slicer is also named **Fruit Slicer** for future reference.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Pivot");
    var slicer = sheet.slicers.add(
        "Farm Sales" /* The slicer data source. For PivotTables, this can be the PivotTable object reference or name. */,
        "Type" /* The field in the data to filter by. For PivotTables, this can be a PivotField object reference or ID. */
    );
    slicer.name = "Fruit Slicer";
    return context.sync();
});
```

#### <a name="filter-items-with-a-slicer"></a><span data-ttu-id="f70b0-264">Filtrar itens com uma segmentação de um</span><span class="sxs-lookup"><span data-stu-id="f70b0-264">Filter items with a slicer</span></span>

<span data-ttu-id="f70b0-265">A segmentação de relatório filtra a tabela dinâmica com itens do `sourceField` .</span><span class="sxs-lookup"><span data-stu-id="f70b0-265">The slicer filters the PivotTable with items from the `sourceField`.</span></span> <span data-ttu-id="f70b0-266">O `Slicer.selectItems` método define os itens que permanecem na segmentação de,.</span><span class="sxs-lookup"><span data-stu-id="f70b0-266">The `Slicer.selectItems` method sets the items that remain in the slicer.</span></span> <span data-ttu-id="f70b0-267">Esses itens são passados para o método como a `string[]` , representando as chaves dos itens.</span><span class="sxs-lookup"><span data-stu-id="f70b0-267">These items are passed to the method as a `string[]`, representing the keys of the items.</span></span> <span data-ttu-id="f70b0-268">Qualquer linha que contenha esses itens permanecerá na agregação da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="f70b0-268">Any rows containing those items remain in the PivotTable's aggregation.</span></span> <span data-ttu-id="f70b0-269">Chamadas subsequentes para `selectItems` definir a lista como as chaves especificadas nessas chamadas.</span><span class="sxs-lookup"><span data-stu-id="f70b0-269">Subsequent calls to `selectItems` set the list to the keys specified in those calls.</span></span>

> [!NOTE]
> <span data-ttu-id="f70b0-270">Se `Slicer.selectItems` for passado um item que não está na fonte de dados, um `InvalidArgument` erro será gerado.</span><span class="sxs-lookup"><span data-stu-id="f70b0-270">If `Slicer.selectItems` is passed an item that's not in the data source, an `InvalidArgument` error is thrown.</span></span> <span data-ttu-id="f70b0-271">O conteúdo pode ser verificado através da `Slicer.slicerItems` propriedade, que é um [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span><span class="sxs-lookup"><span data-stu-id="f70b0-271">The contents can be verified through the `Slicer.slicerItems` property, which is a [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span></span>

<span data-ttu-id="f70b0-272">O exemplo de código a seguir mostra três itens que estão sendo selecionados para a segmentação de itens: **casca** de limão, **verde-limão** e **laranja**.</span><span class="sxs-lookup"><span data-stu-id="f70b0-272">The following code sample shows three items being selected for the slicer: **Lemon**, **Lime**, and **Orange**.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

<span data-ttu-id="f70b0-273">Para remover todos os filtros da segmentação de itens, use o `Slicer.clearFilters` método, conforme mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="f70b0-273">To remove all filters from the slicer, use the `Slicer.clearFilters` method, as shown in the following sample.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a><span data-ttu-id="f70b0-274">Estilo e formatação de uma segmentação de subconjuntos</span><span class="sxs-lookup"><span data-stu-id="f70b0-274">Style and format a slicer</span></span>

<span data-ttu-id="f70b0-275">O suplemento pode ajustar as configurações de exibição de uma segmentação por meio de `Slicer` Propriedades.</span><span class="sxs-lookup"><span data-stu-id="f70b0-275">You add-in can adjust a slicer's display settings through `Slicer` properties.</span></span> <span data-ttu-id="f70b0-276">O exemplo de código a seguir define o estilo como **SlicerStyleLight6**, define o texto na parte superior da segmentação de texto para **tipos de frutas**, coloca a segmentação de texto na posição **(395, 15)** na camada de desenho e define o tamanho da segmentação de texto como **135x150** pixels.</span><span class="sxs-lookup"><span data-stu-id="f70b0-276">The following code sample sets the style to **SlicerStyleLight6**, sets the text at the top of the slicer to **Fruit Types**, places the slicer at the position **(395, 15)** on the drawing layer, and sets the slicer's size to **135x150** pixels.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.caption = "Fruit Types";
    slicer.left = 395;
    slicer.top = 15;
    slicer.height = 135;
    slicer.width = 150;
    slicer.style = "SlicerStyleLight6";
    return context.sync();
});
```

#### <a name="delete-a-slicer"></a><span data-ttu-id="f70b0-277">Excluir uma segmentação de um</span><span class="sxs-lookup"><span data-stu-id="f70b0-277">Delete a slicer</span></span>

<span data-ttu-id="f70b0-278">Para excluir uma segmentação de, chame o `Slicer.delete` método.</span><span class="sxs-lookup"><span data-stu-id="f70b0-278">To delete a slicer, call the `Slicer.delete` method.</span></span> <span data-ttu-id="f70b0-279">O exemplo de código a seguir exclui a primeira segmentação de itens da planilha atual.</span><span class="sxs-lookup"><span data-stu-id="f70b0-279">The following code sample deletes the first slicer from the current worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="f70b0-280">Função de agregação de alteração</span><span class="sxs-lookup"><span data-stu-id="f70b0-280">Change aggregation function</span></span>

<span data-ttu-id="f70b0-281">As hierarquias de dados têm seus valores agregados.</span><span class="sxs-lookup"><span data-stu-id="f70b0-281">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="f70b0-282">Para conjuntos de números de valores, esta é uma soma por padrão.</span><span class="sxs-lookup"><span data-stu-id="f70b0-282">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="f70b0-283">A `summarizeBy` propriedade define esse comportamento com base em um tipo [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) .</span><span class="sxs-lookup"><span data-stu-id="f70b0-283">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="f70b0-284">Os tipos de função de agregação suportados atualmente são,,,,, `Sum` `Count` ,,,, `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` e `Automatic` (o padrão).</span><span class="sxs-lookup"><span data-stu-id="f70b0-284">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="f70b0-285">O exemplo de código a seguir altera a agregação para ser a média dos dados.</span><span class="sxs-lookup"><span data-stu-id="f70b0-285">The following code samples changes the aggregation to be averages of the data.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.dataHierarchies.load("no-properties-needed");
    return context.sync().then(function() {

        // Change the aggregation from the default sum to an average of all the values in the hierarchy.
        pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
        pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
        return context.sync();
    });
});
```

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="f70b0-286">Alterar cálculos com um ShowAsRule</span><span class="sxs-lookup"><span data-stu-id="f70b0-286">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="f70b0-287">As tabelas dinâmicas, por padrão, agregam os dados de suas hierarquias de linha e coluna de forma independente.</span><span class="sxs-lookup"><span data-stu-id="f70b0-287">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="f70b0-288">Um [ShowAsRule](/javascript/api/excel/excel.showasrule) altera a hierarquia de dados para valores de saída com base em outros itens na tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="f70b0-288">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="f70b0-289">O `ShowAsRule` objeto tem três propriedades:</span><span class="sxs-lookup"><span data-stu-id="f70b0-289">The `ShowAsRule` object has three properties:</span></span>

- <span data-ttu-id="f70b0-290">`calculation`: O tipo de cálculo relativo a ser aplicado à hierarquia de dados (o padrão é `none` ).</span><span class="sxs-lookup"><span data-stu-id="f70b0-290">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
- <span data-ttu-id="f70b0-291">`baseField`: O [PivotField](/javascript/api/excel/excel.pivotfield) na hierarquia que contém os dados básicos antes do cálculo ser aplicado.</span><span class="sxs-lookup"><span data-stu-id="f70b0-291">`baseField`: The [PivotField](/javascript/api/excel/excel.pivotfield) in the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="f70b0-292">Como as tabelas dinâmicas do Excel têm um mapeamento de um-para-um de hierarquia para campo, você usará o mesmo nome para acessar a hierarquia e o campo.</span><span class="sxs-lookup"><span data-stu-id="f70b0-292">Since Excel PivotTables have a one-to-one mapping of hierarchy to field, you'll use the same name to access both the hierarchy and the field.</span></span>
- <span data-ttu-id="f70b0-293">`baseItem`: O [PivotItem](/javascript/api/excel/excel.pivotitem) individual comparado com os valores dos campos base com base no tipo de cálculo.</span><span class="sxs-lookup"><span data-stu-id="f70b0-293">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="f70b0-294">Nem todos os cálculos exigem esse campo.</span><span class="sxs-lookup"><span data-stu-id="f70b0-294">Not all calculations require this field.</span></span>

<span data-ttu-id="f70b0-295">O exemplo a seguir define o cálculo **da soma das enações vendidas na** hierarquia de dados do farm como uma porcentagem do total da coluna.</span><span class="sxs-lookup"><span data-stu-id="f70b0-295">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span>
<span data-ttu-id="f70b0-296">Ainda queremos que a granularidade seja estendida para o nível de tipo de frutas, portanto, usaremos a hierarquia de linha de **tipo** e seu campo base.</span><span class="sxs-lookup"><span data-stu-id="f70b0-296">We still want the granularity to extend to the fruit type level, so we'll use the **Type** row hierarchy and its underlying field.</span></span>
<span data-ttu-id="f70b0-297">O exemplo também tem o **farm** como a primeira hierarquia de linha, portanto, o total de entradas do farm exibe a porcentagem de produção de cada farm também.</span><span class="sxs-lookup"><span data-stu-id="f70b0-297">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![Uma tabela dinâmica mostrando as porcentagens das vendas de frutas em relação ao total geral de farms individuais e tipos de frutas individuais em cada farm.](../images/excel-pivots-showas-percentage.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {

        // Show the crates of each fruit type sold at the farm as a percentage of the column's total.
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Percentage of Total Farm Sales";
    });
});
```

<span data-ttu-id="f70b0-299">O exemplo anterior definiu o cálculo para a coluna, em relação ao campo de uma hierarquia de linha individual.</span><span class="sxs-lookup"><span data-stu-id="f70b0-299">The previous example set the calculation to the column, relative to the field of an individual row hierarchy.</span></span> <span data-ttu-id="f70b0-300">Quando o cálculo está relacionado a um item individual, use a `baseItem` propriedade.</span><span class="sxs-lookup"><span data-stu-id="f70b0-300">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="f70b0-301">O exemplo a seguir mostra o `differenceFrom` cálculo.</span><span class="sxs-lookup"><span data-stu-id="f70b0-301">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="f70b0-302">Ele exibe a diferença entre as entradas de hierarquia de dados de vendas do farm em relação às de **um farm**.</span><span class="sxs-lookup"><span data-stu-id="f70b0-302">It displays the difference of the farm crate sales data hierarchy entries relative to those of **A Farms**.</span></span>
<span data-ttu-id="f70b0-303">O `baseField` **farm** de is, portanto, vemos as diferenças entre os outros farms, bem como as divisões de cada tipo de fruta (**Type** também é uma hierarquia de linha neste exemplo).</span><span class="sxs-lookup"><span data-stu-id="f70b0-303">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![Uma tabela dinâmica mostrando as diferenças das vendas de frutas entre "um farm" e outros.](../images/excel-pivots-showas-differencefrom.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {
        // Show the difference between crate sales of the "A Farms" and the other farms.
        // This difference is both aggregated and shown for individual fruit types (where applicable).
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");
        farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Difference from A Farms";
    });
});
```

## <a name="change-hierarchy-names"></a><span data-ttu-id="f70b0-307">Alterar nomes de hierarquia</span><span class="sxs-lookup"><span data-stu-id="f70b0-307">Change hierarchy names</span></span>

<span data-ttu-id="f70b0-308">Os campos de hierarquia são editáveis.</span><span class="sxs-lookup"><span data-stu-id="f70b0-308">Hierarchy fields are editable.</span></span> <span data-ttu-id="f70b0-309">O código a seguir demonstra como alterar os nomes exibidos de duas hierarquias de dados.</span><span class="sxs-lookup"><span data-stu-id="f70b0-309">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

```js
Excel.run(function (context) {
    var dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    return context.sync().then(function () {
        // changing the displayed names of these entries
        dataHierarchies.items[0].name = "Farm Sales";
        dataHierarchies.items[1].name = "Wholesale";
    });
});
```

## <a name="see-also"></a><span data-ttu-id="f70b0-310">Confira também</span><span class="sxs-lookup"><span data-stu-id="f70b0-310">See also</span></span>

- [<span data-ttu-id="f70b0-311">Modelo de objeto do JavaScript do Excel em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="f70b0-311">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="f70b0-312">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f70b0-312">Excel JavaScript API Reference</span></span>](/javascript/api/excel)
