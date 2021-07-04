---
title: Trabalhar com tabelas dinâmicas usando a Excel JavaScript
description: Use a Excel JavaScript para criar tabelas dinâmicas e interagir com seus componentes.
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: 8c8917f57b7546694e12380fc4369847be24ceac
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290737"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="630df-103">Trabalhar com tabelas dinâmicas usando a Excel JavaScript</span><span class="sxs-lookup"><span data-stu-id="630df-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="630df-104">Tabelas dinâmicas simplificam conjuntos de dados maiores.</span><span class="sxs-lookup"><span data-stu-id="630df-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="630df-105">Eles permitem a manipulação rápida de dados agrupados.</span><span class="sxs-lookup"><span data-stu-id="630df-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="630df-106">A Excel API JavaScript permite que seu complemento crie Tabelas Dinâmicas e interaja com seus componentes.</span><span class="sxs-lookup"><span data-stu-id="630df-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> <span data-ttu-id="630df-107">Este artigo descreve como as Tabelas Dinâmicas são representadas pela API JavaScript Office e fornece exemplos de código para cenários principais.</span><span class="sxs-lookup"><span data-stu-id="630df-107">This article describes how PivotTables are represented by the Office JavaScript API and provides code samples for key scenarios.</span></span>

<span data-ttu-id="630df-108">Se você não estiver familiarizado com a funcionalidade das Tabelas Dinâmicas, considere explorá-las como um usuário final.</span><span class="sxs-lookup"><span data-stu-id="630df-108">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span>
<span data-ttu-id="630df-109">Consulte [Criar uma Tabela Dinâmica para analisar dados de planilha](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) para uma boa cartilha nessas ferramentas.</span><span class="sxs-lookup"><span data-stu-id="630df-109">See [Create a PivotTable to analyze worksheet data](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="630df-110">As Tabelas Dinâmicas criadas com o OLAP não são suportadas no momento.</span><span class="sxs-lookup"><span data-stu-id="630df-110">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="630df-111">Também não há suporte para o Power Pivot.</span><span class="sxs-lookup"><span data-stu-id="630df-111">There is also no support for Power Pivot.</span></span>

## <a name="object-model"></a><span data-ttu-id="630df-112">Modelo de objetos</span><span class="sxs-lookup"><span data-stu-id="630df-112">Object model</span></span>

<span data-ttu-id="630df-113">A [Tabela Dinâmica](/javascript/api/excel/excel.pivottable) é o objeto central para Tabelas Dinâmicas na API JavaScript Office JavaScript.</span><span class="sxs-lookup"><span data-stu-id="630df-113">The [PivotTable](/javascript/api/excel/excel.pivottable) is the central object for PivotTables in the Office JavaScript API.</span></span>

- <span data-ttu-id="630df-114">`Workbook.pivotTables` e `Worksheet.pivotTables` são [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) que contêm as [Tabelas Dinâmicas](/javascript/api/excel/excel.pivottable) na pasta de trabalho e planilha, respectivamente.</span><span class="sxs-lookup"><span data-stu-id="630df-114">`Workbook.pivotTables` and `Worksheet.pivotTables` are [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that contain the [PivotTables](/javascript/api/excel/excel.pivottable) in the workbook and worksheet, respectively.</span></span>
- <span data-ttu-id="630df-115">Uma [Tabela Dinâmica](/javascript/api/excel/excel.pivottable) contém um [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) que tem vários [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span><span class="sxs-lookup"><span data-stu-id="630df-115">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) that has multiple [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span></span>
- <span data-ttu-id="630df-116">Essas [pivotHierarchies](/javascript/api/excel/excel.pivothierarchy) podem ser adicionadas a coleções de hierarquia específicas para definir como os dados de pivot de tabela dinâmica (conforme explicado na [seção a seguir](#hierarchies)).</span><span class="sxs-lookup"><span data-stu-id="630df-116">These [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) can be added to specific hierarchy collections to define how the PivotTable pivots data (as explained in the [following section](#hierarchies)).</span></span>
- <span data-ttu-id="630df-117">Um [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contém [um PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) que tem exatamente um [PivotField](/javascript/api/excel/excel.pivotfield).</span><span class="sxs-lookup"><span data-stu-id="630df-117">A [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contains a [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) that has exactly one [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="630df-118">Se o design se expandir para incluir tabelas dinâmicas OLAP, isso poderá mudar.</span><span class="sxs-lookup"><span data-stu-id="630df-118">If the design expands to include OLAP PivotTables, this may change.</span></span>
- <span data-ttu-id="630df-119">Um [PivotField](/javascript/api/excel/excel.pivotfield) pode ter um ou mais [PivotFilters aplicados,](/javascript/api/excel/excel.pivotfilters) desde que [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) do campo seja atribuído a uma categoria de hierarquia.</span><span class="sxs-lookup"><span data-stu-id="630df-119">A [PivotField](/javascript/api/excel/excel.pivotfield) can have one or more [PivotFilters](/javascript/api/excel/excel.pivotfilters) applied, as long as the field's [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) is assigned to a hierarchy category.</span></span>
- <span data-ttu-id="630df-120">Um [PivotField](/javascript/api/excel/excel.pivotfield) contém um [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) que tem vários [PivotItems](/javascript/api/excel/excel.pivotitem).</span><span class="sxs-lookup"><span data-stu-id="630df-120">A [PivotField](/javascript/api/excel/excel.pivotfield) contains a [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) that has multiple [PivotItems](/javascript/api/excel/excel.pivotitem).</span></span>
- <span data-ttu-id="630df-121">Uma [Tabela Dinâmica](/javascript/api/excel/excel.pivottable) contém um [PivotLayout](/javascript/api/excel/excel.pivotlayout) que define onde os [PivotFields](/javascript/api/excel/excel.pivotfield) e [PivotItems](/javascript/api/excel/excel.pivotitem) são exibidos na planilha.</span><span class="sxs-lookup"><span data-stu-id="630df-121">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotLayout](/javascript/api/excel/excel.pivotlayout) that defines where the [PivotFields](/javascript/api/excel/excel.pivotfield) and [PivotItems](/javascript/api/excel/excel.pivotitem) are displayed in the worksheet.</span></span> <span data-ttu-id="630df-122">O layout também controla algumas configurações de exibição para a Tabela Dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-122">The layout also controls some display settings for the PivotTable.</span></span>

<span data-ttu-id="630df-123">Vejamos como essas relações se aplicam a alguns dados de exemplo.</span><span class="sxs-lookup"><span data-stu-id="630df-123">Let's look at how these relationships apply to some example data.</span></span> <span data-ttu-id="630df-124">Os dados a seguir descrevem as vendas de frutas de vários farms.</span><span class="sxs-lookup"><span data-stu-id="630df-124">The following data describes fruit sales from various farms.</span></span> <span data-ttu-id="630df-125">Ele será o exemplo ao longo deste artigo.</span><span class="sxs-lookup"><span data-stu-id="630df-125">It will be the example throughout this article.</span></span>

![Uma coleção de vendas de frutas de diferentes tipos de farms diferentes.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="630df-127">Esses dados de vendas de farm de frutas serão usados para fazer uma tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-127">This fruit farm sales data will be used to make a PivotTable.</span></span> <span data-ttu-id="630df-128">Cada coluna, como **Tipos,** é `PivotHierarchy` um .</span><span class="sxs-lookup"><span data-stu-id="630df-128">Each column, such as **Types**, is a `PivotHierarchy`.</span></span> <span data-ttu-id="630df-129">A **hierarquia Tipos** contém o campo **Tipos.**</span><span class="sxs-lookup"><span data-stu-id="630df-129">The **Types** hierarchy contains the **Types** field.</span></span> <span data-ttu-id="630df-130">O **campo Tipos** contém os itens **Apple,** **Kiwi,** **Limão,** **Lima** e **Laranja**.</span><span class="sxs-lookup"><span data-stu-id="630df-130">The **Types** field contains the items **Apple**, **Kiwi**, **Lemon**, **Lime**, and **Orange**.</span></span>

### <a name="hierarchies"></a><span data-ttu-id="630df-131">Hierarquias</span><span class="sxs-lookup"><span data-stu-id="630df-131">Hierarchies</span></span>

<span data-ttu-id="630df-132">As Tabelas Dinâmicas são organizadas com base em quatro categorias de hierarquia: [linha,](/javascript/api/excel/excel.rowcolumnpivothierarchy) [coluna,](/javascript/api/excel/excel.rowcolumnpivothierarchy) [dados](/javascript/api/excel/excel.datapivothierarchy)e [filtro](/javascript/api/excel/excel.filterpivothierarchy).</span><span class="sxs-lookup"><span data-stu-id="630df-132">PivotTables are organized based on four hierarchy categories: [row](/javascript/api/excel/excel.rowcolumnpivothierarchy), [column](/javascript/api/excel/excel.rowcolumnpivothierarchy), [data](/javascript/api/excel/excel.datapivothierarchy), and [filter](/javascript/api/excel/excel.filterpivothierarchy).</span></span>

<span data-ttu-id="630df-133">Os dados do farm mostrados anteriormente têm cinco hierarquias: **Farms**, **Type**, **Classification,** **Crates Sold at Farm** e **Crates Sold Fim de Semana.**</span><span class="sxs-lookup"><span data-stu-id="630df-133">The farm data shown earlier has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="630df-134">Cada hierarquia só pode existir em uma das quatro categorias.</span><span class="sxs-lookup"><span data-stu-id="630df-134">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="630df-135">Se **Type** for adicionado a hierarquias de coluna, ele também não poderá estar na linha, dados ou hierarquias de filtro.</span><span class="sxs-lookup"><span data-stu-id="630df-135">If **Type** is added to column hierarchies, it cannot also be in the row, data, or filter hierarchies.</span></span> <span data-ttu-id="630df-136">Se **Type** for subsequentemente adicionado às hierarquias de linha, ele será removido das hierarquias de coluna.</span><span class="sxs-lookup"><span data-stu-id="630df-136">If **Type** is subsequently added to row hierarchies, it is removed from the column hierarchies.</span></span> <span data-ttu-id="630df-137">Esse comportamento é o mesmo se a atribuição de hierarquia é feita por meio da interface do usuário Excel ou do Excel APIs JavaScript.</span><span class="sxs-lookup"><span data-stu-id="630df-137">This behavior is the same whether hierarchy assignment is done through the Excel UI or the Excel JavaScript APIs.</span></span>

<span data-ttu-id="630df-138">Hierarquias de linhas e colunas definem como os dados serão agrupados.</span><span class="sxs-lookup"><span data-stu-id="630df-138">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="630df-139">Por exemplo, uma hierarquia de linhas **de Farms** agrupa todos os conjuntos de dados do mesmo farm.</span><span class="sxs-lookup"><span data-stu-id="630df-139">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="630df-140">A escolha entre a hierarquia de linha e coluna define a orientação da Tabela Dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-140">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="630df-141">Hierarquias de dados são os valores a serem agregados com base nas hierarquias de linha e coluna.</span><span class="sxs-lookup"><span data-stu-id="630df-141">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="630df-142">Uma Tabela Dinâmica com uma hierarquia de linhas de **Farms** e uma hierarquia de dados de **Engradados Vendidos por** Atacado mostra a soma total (por padrão) de todas as diferentes frutas para cada farm.</span><span class="sxs-lookup"><span data-stu-id="630df-142">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="630df-143">Hierarquias de filtro incluem ou excluem dados do pivô com base nos valores dentro desse tipo filtrado.</span><span class="sxs-lookup"><span data-stu-id="630df-143">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="630df-144">Uma hierarquia de filtro de **Classificação com** o tipo **Organic** selecionado mostra apenas dados para frutas orgânicas.</span><span class="sxs-lookup"><span data-stu-id="630df-144">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="630df-145">Aqui estão os dados do farm novamente, juntamente com uma tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-145">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="630df-146">A Tabela Dinâmica  está usando **Farm** e **Type** como hierarquias de linha, Caixas **Vendidas** no Farm e Engradados Vendidos por Atacado como **hierarquias** de dados (com a função de agregação padrão de soma) e Classificação como uma hierarquia de filtro (com a **seleção** orgânica).</span><span class="sxs-lookup"><span data-stu-id="630df-146">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span>

![Uma seleção de dados de vendas de frutas ao lado de uma tabela dinâmica com hierarquias de linha, dados e filtro.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="630df-148">Essa tabela dinâmica pode ser gerada por meio da API JavaScript ou por meio da interface do usuário Excel usuário.</span><span class="sxs-lookup"><span data-stu-id="630df-148">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="630df-149">Ambas as opções permitem mais manipulação por meio de complementos.</span><span class="sxs-lookup"><span data-stu-id="630df-149">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="630df-150">Criar uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="630df-150">Create a PivotTable</span></span>

<span data-ttu-id="630df-151">As Tabelas Dinâmicas precisam de um nome, fonte e destino.</span><span class="sxs-lookup"><span data-stu-id="630df-151">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="630df-152">A origem pode ser um endereço de intervalo ou um nome de tabela (passado como `Range` `string` , ou `Table` tipo).</span><span class="sxs-lookup"><span data-stu-id="630df-152">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="630df-153">O destino é um endereço de intervalo (dado como a `Range` ou `string` ).</span><span class="sxs-lookup"><span data-stu-id="630df-153">The destination is a range address (given as either a `Range` or `string`).</span></span>
<span data-ttu-id="630df-154">Os exemplos a seguir mostram várias técnicas de criação de tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-154">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="630df-155">Criar uma tabela dinâmica com endereços de intervalo</span><span class="sxs-lookup"><span data-stu-id="630df-155">Create a PivotTable with range addresses</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="630df-156">Criar uma tabela dinâmica com objetos Range</span><span class="sxs-lookup"><span data-stu-id="630df-156">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="630df-157">Criar uma Tabela Dinâmica no nível da workbook</span><span class="sxs-lookup"><span data-stu-id="630df-157">Create a PivotTable at the workbook level</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="630df-158">Usar uma tabela dinâmica existente</span><span class="sxs-lookup"><span data-stu-id="630df-158">Use an existing PivotTable</span></span>

<span data-ttu-id="630df-159">As Tabelas Dinâmicas criadas manualmente também são acessíveis por meio da coleção PivotTable da pasta de trabalho ou de planilhas individuais.</span><span class="sxs-lookup"><span data-stu-id="630df-159">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> <span data-ttu-id="630df-160">O código a seguir obtém uma Tabela Dinâmica chamada **Meu Pivô** da lista de trabalho.</span><span class="sxs-lookup"><span data-stu-id="630df-160">The following code gets a PivotTable named **My Pivot** from the workbook.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="630df-161">Adicionar linhas e colunas a uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="630df-161">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="630df-162">Linhas e colunas giram os dados em torno dos valores desses campos.</span><span class="sxs-lookup"><span data-stu-id="630df-162">Rows and columns pivot the data around those fields' values.</span></span>

<span data-ttu-id="630df-163">A **adição da coluna Farm** gira todas as vendas ao redor de cada farm.</span><span class="sxs-lookup"><span data-stu-id="630df-163">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="630df-164">Adicionar as **linhas Tipo** e **Classificação** quebra ainda mais os dados com base em quais frutas foram vendidas e se foram orgânicas ou não.</span><span class="sxs-lookup"><span data-stu-id="630df-164">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

![Uma tabela dinâmica com uma coluna farm e linhas Tipo e Classificação.](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

<span data-ttu-id="630df-166">Você também pode ter uma Tabela Dinâmica com apenas linhas ou colunas.</span><span class="sxs-lookup"><span data-stu-id="630df-166">You can also have a PivotTable with only rows or columns.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="630df-167">Adicionar hierarquias de dados à Tabela Dinâmica</span><span class="sxs-lookup"><span data-stu-id="630df-167">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="630df-168">Hierarquias de dados preenchem a Tabela Dinâmica com informações para combinar com base nas linhas e colunas.</span><span class="sxs-lookup"><span data-stu-id="630df-168">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="630df-169">A adição das hierarquias de dados de **Caixas Vendidas** no Farm e caixas **vendidas por** atacado fornece somas desses números para cada linha e coluna.</span><span class="sxs-lookup"><span data-stu-id="630df-169">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span>

<span data-ttu-id="630df-170">No exemplo, **Farm** e **Type** são linhas, com as vendas do engradado como os dados.</span><span class="sxs-lookup"><span data-stu-id="630df-170">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span>

![Uma Tabela Dinâmica mostrando o total de vendas de diferentes frutas com base no farm de onde eles vieram.](../images/excel-pivots-data-hierarchy.png)

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

## <a name="pivottable-layouts-and-getting-pivoted-data"></a><span data-ttu-id="630df-172">Layouts de tabela dinâmica e informações dinâmicas</span><span class="sxs-lookup"><span data-stu-id="630df-172">PivotTable layouts and getting pivoted data</span></span>

<span data-ttu-id="630df-173">Um [PivotLayout](/javascript/api/excel/excel.pivotlayout) define o posicionamento das hierarquias e seus dados.</span><span class="sxs-lookup"><span data-stu-id="630df-173">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="630df-174">Você acessa o layout para determinar os intervalos onde os dados são armazenados.</span><span class="sxs-lookup"><span data-stu-id="630df-174">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="630df-175">O diagrama a seguir mostra quais chamadas de função de layout correspondem a quais intervalos da Tabela Dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-175">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Um diagrama mostrando quais seções de uma tabela dinâmica são retornadas pelas funções de intervalo de obter do layout.](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a><span data-ttu-id="630df-177">Obter dados da tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="630df-177">Get data from the PivotTable</span></span>

<span data-ttu-id="630df-178">O layout define como a Tabela Dinâmica é exibida na planilha.</span><span class="sxs-lookup"><span data-stu-id="630df-178">The layout defines how the PivotTable is displayed in the worksheet.</span></span> <span data-ttu-id="630df-179">Isso significa que `PivotLayout` o objeto controla os intervalos usados para elementos de tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-179">This means the `PivotLayout` object controls the ranges used for PivotTable elements.</span></span> <span data-ttu-id="630df-180">Use os intervalos fornecidos pelo layout para obter dados coletados e agregados pela Tabela Dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-180">Use the ranges provided by the layout to get data collected and aggregated by the PivotTable.</span></span> <span data-ttu-id="630df-181">Em particular, use `PivotLayout.getDataBodyRange` para acessar os dados produzidos pela Tabela Dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-181">In particular, use `PivotLayout.getDataBodyRange` to access the data produced by the PivotTable.</span></span>

<span data-ttu-id="630df-182">O código a seguir demonstra como obter a última linha dos dados de tabela dinâmica passando pelo layout (o **Grande Total** da Soma de Caixas **Vendidas** no Farm e a Soma das **colunas De Engradados Vendidos** por Atacado no exemplo anterior).</span><span class="sxs-lookup"><span data-stu-id="630df-182">The following code demonstrates how to get the last row of the PivotTable data by going through the layout (the **Grand Total** of both the **Sum of Crates Sold at Farm** and **Sum of Crates Sold Wholesale** columns in the earlier example).</span></span> <span data-ttu-id="630df-183">Esses valores são, em seguida, resumidos para um total final, que é exibido na célula **E30** (fora da Tabela Dinâmica).</span><span class="sxs-lookup"><span data-stu-id="630df-183">Those values are then summed together for a final total, which is displayed in cell **E30** (outside of the PivotTable).</span></span>

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

### <a name="layout-types"></a><span data-ttu-id="630df-184">Tipos de layout</span><span class="sxs-lookup"><span data-stu-id="630df-184">Layout types</span></span>

<span data-ttu-id="630df-185">As Tabelas Dinâmicas têm três estilos de layout: Compact, Outline e Tabular.</span><span class="sxs-lookup"><span data-stu-id="630df-185">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="630df-186">Vimos o estilo compacto nos exemplos anteriores.</span><span class="sxs-lookup"><span data-stu-id="630df-186">We've seen the compact style in the previous examples.</span></span>

<span data-ttu-id="630df-187">Os exemplos a seguir usam os estilos de contorno e tabular, respectivamente.</span><span class="sxs-lookup"><span data-stu-id="630df-187">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="630df-188">O exemplo de código mostra como fazer o ciclo entre os diferentes layouts.</span><span class="sxs-lookup"><span data-stu-id="630df-188">The code sample shows how to cycle between the different layouts.</span></span>

#### <a name="outline-layout"></a><span data-ttu-id="630df-189">Layout de estrutura de estrutura</span><span class="sxs-lookup"><span data-stu-id="630df-189">Outline layout</span></span>

![Uma tabela dinâmica usando o layout de estrutura de estrutura.](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a><span data-ttu-id="630df-191">Layout tabular</span><span class="sxs-lookup"><span data-stu-id="630df-191">Tabular layout</span></span>

![Uma tabela dinâmica usando o layout tabular.](../images/excel-pivots-tabular-layout.png)

#### <a name="pivotlayout-type-switch-code-sample"></a><span data-ttu-id="630df-193">Exemplo de código de opção do tipo PivotLayout</span><span class="sxs-lookup"><span data-stu-id="630df-193">PivotLayout type switch code sample</span></span>

```js
Excel.run(function (context) {
    // Change the PivotLayout.type to a new type.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.layout.load("layoutType");
    return context.sync().then(function () {
        // Cycle between the three layout types.
        if (pivotTable.layout.layoutType === "Compact") {
            pivotTable.layout.layoutType = "Outline";
        } else if (pivotTable.layout.layoutType === "Outline") {
            pivotTable.layout.layoutType = "Tabular";
        } else {
            pivotTable.layout.layoutType = "Compact";
        }
    
        return context.sync();
    });
});
```

### <a name="other-pivotlayout-functions"></a><span data-ttu-id="630df-194">Outras funções PivotLayout</span><span class="sxs-lookup"><span data-stu-id="630df-194">Other PivotLayout functions</span></span>

<span data-ttu-id="630df-195">Por padrão, tabelas dinâmicas ajustam tamanhos de linha e coluna conforme necessário.</span><span class="sxs-lookup"><span data-stu-id="630df-195">By default, PivotTables adjust row and column sizes as needed.</span></span> <span data-ttu-id="630df-196">Isso é feito quando a Tabela Dinâmica é atualizada.</span><span class="sxs-lookup"><span data-stu-id="630df-196">This is done when the PivotTable is refreshed.</span></span> <span data-ttu-id="630df-197">`PivotLayout.autoFormat` especifica esse comportamento.</span><span class="sxs-lookup"><span data-stu-id="630df-197">`PivotLayout.autoFormat` specifies that behavior.</span></span> <span data-ttu-id="630df-198">Qualquer alteração de tamanho de linha ou coluna feita pelo seu complemento persiste quando `autoFormat` é `false` .</span><span class="sxs-lookup"><span data-stu-id="630df-198">Any row or column size changes made by your add-in persist when `autoFormat` is `false`.</span></span> <span data-ttu-id="630df-199">Além disso, as configurações padrão de uma tabela dinâmica mantêm qualquer formatação personalizada na Tabela Dinâmica (como preenchimentos e alterações de fonte).</span><span class="sxs-lookup"><span data-stu-id="630df-199">Additionally, the default settings of a PivotTable keep any custom formatting in the PivotTable (such as fills and font changes).</span></span> <span data-ttu-id="630df-200">Definir `PivotLayout.preserveFormatting` para aplicar o formato padrão quando `false` atualizado.</span><span class="sxs-lookup"><span data-stu-id="630df-200">Set `PivotLayout.preserveFormatting` to `false` to apply the default format when refreshed.</span></span>

<span data-ttu-id="630df-201">Um também controla as configurações de header e de linha total, como as células de dados vazias são `PivotLayout` exibidas e as opções de texto [alt.](https://support.microsoft.com/topic/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669)</span><span class="sxs-lookup"><span data-stu-id="630df-201">A `PivotLayout` also controls header and total row settings, how empty data cells are displayed, and [alt text](https://support.microsoft.com/topic/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669) options.</span></span> <span data-ttu-id="630df-202">A [referência PivotLayout](/javascript/api/excel/excel.pivotlayout) fornece uma lista completa desses recursos.</span><span class="sxs-lookup"><span data-stu-id="630df-202">The [PivotLayout](/javascript/api/excel/excel.pivotlayout) reference provides a complete list of these features.</span></span>

<span data-ttu-id="630df-203">O exemplo de código a seguir faz com que as células de dados vazias exibem a cadeia de caracteres , formate o intervalo do corpo para um alinhamento horizontal consistente e garante que as alterações de formatação permaneçam mesmo após a atualização da Tabela `"--"` Dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-203">The following code sample makes empty data cells display the string `"--"`, formats the body range to a consistent horizontal alignment, and ensures that the formatting changes remain even after the PivotTable is refreshed.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("Farm Sales");
    var pivotLayout = pivotTable.layout;

    // Set a default value for an empty cell in the PivotTable. This doesn't include cells left blank by the layout.
    pivotLayout.emptyCellText = "--";

    // Set the text alignment to match the rest of the PivotTable.
    pivotLayout.getDataBodyRange().format.horizontalAlignment = Excel.HorizontalAlignment.right;

    // Ensure empty cells are filled with a default value.
    pivotLayout.fillEmptyCells = true;

    // Ensure that the format settings persist, even after the PivotTable is refreshed and recalculated.
    pivotLayout.preserveFormatting = true;
    return context.sync();
});
```

## <a name="delete-a-pivottable"></a><span data-ttu-id="630df-204">Excluir uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="630df-204">Delete a PivotTable</span></span>

<span data-ttu-id="630df-205">Tabelas Dinâmicas são excluídas usando seu nome.</span><span class="sxs-lookup"><span data-stu-id="630df-205">PivotTables are deleted by using their name.</span></span>

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a><span data-ttu-id="630df-206">Filtrar uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="630df-206">Filter a PivotTable</span></span>

<span data-ttu-id="630df-207">O método principal para filtrar dados de tabela dinâmica é com PivotFilters.</span><span class="sxs-lookup"><span data-stu-id="630df-207">The primary method for filtering PivotTable data is with PivotFilters.</span></span> <span data-ttu-id="630df-208">As slicers oferecem um método alternativo de filtragem menos flexível.</span><span class="sxs-lookup"><span data-stu-id="630df-208">Slicers offer an alternate, less flexible filtering method.</span></span>

<span data-ttu-id="630df-209">[PivotFilters](/javascript/api/excel/excel.pivotfilters) filtram dados com base [](#hierarchies) nas quatro categorias de hierarquia de uma tabela dinâmica (filtros, colunas, linhas e valores).</span><span class="sxs-lookup"><span data-stu-id="630df-209">[PivotFilters](/javascript/api/excel/excel.pivotfilters) filter data based on a PivotTable's four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="630df-210">Há quatro tipos de PivotFilters, permitindo filtragem baseada em data de calendário, análise de cadeia de caracteres, comparação de números e filtragem com base em uma entrada personalizada.</span><span class="sxs-lookup"><span data-stu-id="630df-210">There are four types of PivotFilters, allowing calendar date-based filtering, string parsing, number comparison, and filtering based on a custom input.</span></span>

<span data-ttu-id="630df-211">[As slicers](/javascript/api/excel/excel.slicer) podem ser aplicadas a tabelas dinâmicas e Excel regulares.</span><span class="sxs-lookup"><span data-stu-id="630df-211">[Slicers](/javascript/api/excel/excel.slicer) can be applied to both PivotTables and regular Excel tables.</span></span> <span data-ttu-id="630df-212">Quando aplicada a uma Tabela Dinâmica, as slicers funcionam como um [PivotManualFilter](#pivotmanualfilter) e permitem a filtragem com base em uma entrada personalizada.</span><span class="sxs-lookup"><span data-stu-id="630df-212">When applied to a PivotTable, slicers function like a [PivotManualFilter](#pivotmanualfilter) and allow filtering based on a custom input.</span></span> <span data-ttu-id="630df-213">Ao contrário dos PivotFilters, as slicers têm um [Excel de interface do usuário](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d).</span><span class="sxs-lookup"><span data-stu-id="630df-213">Unlike PivotFilters, slicers have an [Excel UI component](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d).</span></span> <span data-ttu-id="630df-214">Com a `Slicer` classe, você cria esse componente de interface do usuário, gerencia a filtragem e controla sua aparência visual.</span><span class="sxs-lookup"><span data-stu-id="630df-214">With the `Slicer` class, you create this UI component, manage filtering, and control its visual appearance.</span></span>

### <a name="filter-with-pivotfilters"></a><span data-ttu-id="630df-215">Filtrar com PivotFilters</span><span class="sxs-lookup"><span data-stu-id="630df-215">Filter with PivotFilters</span></span>

<span data-ttu-id="630df-216">[Os PivotFilters](/javascript/api/excel/excel.pivotfilters) permitem filtrar dados [](#hierarchies) de tabela dinâmica com base nas quatro categorias de hierarquia (filtros, colunas, linhas e valores).</span><span class="sxs-lookup"><span data-stu-id="630df-216">[PivotFilters](/javascript/api/excel/excel.pivotfilters) allow you to filter PivotTable data based on the four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="630df-217">No modelo de objeto pivotTable, `PivotFilters` são aplicados a um [PivotField](/javascript/api/excel/excel.pivotfield), e cada um pode `PivotField` ter um ou mais `PivotFilters` atribuídos .</span><span class="sxs-lookup"><span data-stu-id="630df-217">In the PivotTable object model, `PivotFilters` are applied to a [PivotField](/javascript/api/excel/excel.pivotfield), and each `PivotField` can have one or more assigned `PivotFilters`.</span></span> <span data-ttu-id="630df-218">Para aplicar PivotFilters a um PivotField, a [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) correspondente do campo deve ser atribuída a uma categoria de hierarquia.</span><span class="sxs-lookup"><span data-stu-id="630df-218">To apply PivotFilters to a PivotField, the field's corresponding [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) must be assigned to a hierarchy category.</span></span>

#### <a name="types-of-pivotfilters"></a><span data-ttu-id="630df-219">Tipos de PivotFilters</span><span class="sxs-lookup"><span data-stu-id="630df-219">Types of PivotFilters</span></span>

| <span data-ttu-id="630df-220">Tipo de filtro</span><span class="sxs-lookup"><span data-stu-id="630df-220">Filter type</span></span> | <span data-ttu-id="630df-221">Finalidade de filtro</span><span class="sxs-lookup"><span data-stu-id="630df-221">Filter purpose</span></span> | <span data-ttu-id="630df-222">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="630df-222">Excel JavaScript API reference</span></span> |
|:--- |:--- |:--- |
| <span data-ttu-id="630df-223">DateFilter</span><span class="sxs-lookup"><span data-stu-id="630df-223">DateFilter</span></span> | <span data-ttu-id="630df-224">Filtragem baseada em data de calendário.</span><span class="sxs-lookup"><span data-stu-id="630df-224">Calendar date-based filtering.</span></span> | [<span data-ttu-id="630df-225">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="630df-225">PivotDateFilter</span></span>](/javascript/api/excel/excel.pivotdatefilter) |
| <span data-ttu-id="630df-226">LabelFilter</span><span class="sxs-lookup"><span data-stu-id="630df-226">LabelFilter</span></span> | <span data-ttu-id="630df-227">Filtragem de comparação de texto.</span><span class="sxs-lookup"><span data-stu-id="630df-227">Text comparison filtering.</span></span> | [<span data-ttu-id="630df-228">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="630df-228">PivotLabelFilter</span></span>](/javascript/api/excel/excel.pivotlabelfilter) |
| <span data-ttu-id="630df-229">ManualFilter</span><span class="sxs-lookup"><span data-stu-id="630df-229">ManualFilter</span></span> | <span data-ttu-id="630df-230">Filtragem de entrada personalizada.</span><span class="sxs-lookup"><span data-stu-id="630df-230">Custom input filtering.</span></span> | [<span data-ttu-id="630df-231">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="630df-231">PivotManualFilter</span></span>](/javascript/api/excel/excel.pivotmanualfilter) |
| <span data-ttu-id="630df-232">ValueFilter</span><span class="sxs-lookup"><span data-stu-id="630df-232">ValueFilter</span></span> | <span data-ttu-id="630df-233">Filtragem de comparação de números.</span><span class="sxs-lookup"><span data-stu-id="630df-233">Number comparison filtering.</span></span> | [<span data-ttu-id="630df-234">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="630df-234">PivotValueFilter</span></span>](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a><span data-ttu-id="630df-235">Criar um PivotFilter</span><span class="sxs-lookup"><span data-stu-id="630df-235">Create a PivotFilter</span></span>

<span data-ttu-id="630df-236">Para filtrar dados de tabela dinâmica com um (como um ), aplique o `Pivot*Filter` filtro a um `PivotDateFilter` [PivotField](/javascript/api/excel/excel.pivotfield).</span><span class="sxs-lookup"><span data-stu-id="630df-236">To filter PivotTable data with a `Pivot*Filter` (such as a `PivotDateFilter`), apply the filter to a [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="630df-237">Os quatro exemplos de código a seguir mostram como usar cada um dos quatro tipos de PivotFilters.</span><span class="sxs-lookup"><span data-stu-id="630df-237">The following four code samples show how to use each of the four types of PivotFilters.</span></span>

##### <a name="pivotdatefilter"></a><span data-ttu-id="630df-238">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="630df-238">PivotDateFilter</span></span>

<span data-ttu-id="630df-239">O primeiro exemplo de código aplica  um [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) ao PivotField Atualizado de Data, ocultando todos os dados antes **de 2020-08-01**.</span><span class="sxs-lookup"><span data-stu-id="630df-239">The first code sample applies a [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) to the **Date Updated** PivotField, hiding any data prior to **2020-08-01**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="630df-240">Um `Pivot*Filter` não pode ser aplicado a um PivotField, a menos que PivotHierarchy desse campo seja atribuída a uma categoria de hierarquia.</span><span class="sxs-lookup"><span data-stu-id="630df-240">A `Pivot*Filter` can't be applied to a PivotField unless that field's PivotHierarchy is assigned to a hierarchy category.</span></span> <span data-ttu-id="630df-241">No exemplo de código a seguir, o deve ser adicionado à categoria da tabela dinâmica antes de poder `dateHierarchy` `rowHierarchies` ser usado para filtragem.</span><span class="sxs-lookup"><span data-stu-id="630df-241">In the following code sample, the `dateHierarchy` must be added to the PivotTable's `rowHierarchies` category before it can be used for filtering.</span></span>

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
> <span data-ttu-id="630df-242">Os três trechos de código a seguir exibem apenas trechos específicos do filtro, em vez de chamadas `Excel.run` completas.</span><span class="sxs-lookup"><span data-stu-id="630df-242">The following three code snippets only display filter-specific excerpts, instead of full `Excel.run` calls.</span></span>

##### <a name="pivotlabelfilter"></a><span data-ttu-id="630df-243">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="630df-243">PivotLabelFilter</span></span>

<span data-ttu-id="630df-244">O segundo trecho de código demonstra como aplicar um [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) ao **Tipo** PivotField, usando a propriedade para excluir rótulos que começam com a `LabelFilterCondition.beginsWith` letra **L**.</span><span class="sxs-lookup"><span data-stu-id="630df-244">The second code snippet demonstrates how to apply a [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) to the **Type** PivotField, using the `LabelFilterCondition.beginsWith` property to exclude labels that start with the letter **L**.</span></span>

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

##### <a name="pivotmanualfilter"></a><span data-ttu-id="630df-245">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="630df-245">PivotManualFilter</span></span>

<span data-ttu-id="630df-246">O terceiro trecho de código aplica um filtro manual com  [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) ao campo Classificação, filtrando dados que não incluem a classificação **Orgânica**.</span><span class="sxs-lookup"><span data-stu-id="630df-246">The third code snippet applies a manual filter with [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) to the the **Classification** field, filtering out data that doesn't include the classification **Organic**.</span></span>

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a><span data-ttu-id="630df-247">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="630df-247">PivotValueFilter</span></span>

<span data-ttu-id="630df-248">Para comparar números, use um filtro de valor com [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter), conforme mostrado no trecho de código final.</span><span class="sxs-lookup"><span data-stu-id="630df-248">To compare numbers, use a value filter with [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter), as shown in the final code snippet.</span></span> <span data-ttu-id="630df-249">O compara os dados no Farm PivotField com os dados no Campo Pivô de Engradados Vendidos por Atacado, incluindo apenas farms cuja soma de caixas vendidas excede o valor `PivotValueFilter` **de 500**.  </span><span class="sxs-lookup"><span data-stu-id="630df-249">The `PivotValueFilter` compares the data in the **Farm** PivotField to the data in the **Crates Sold Wholesale** PivotField, including only farms whose sum of crates sold exceeds the value **500**.</span></span>

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

#### <a name="remove-pivotfilters"></a><span data-ttu-id="630df-250">Remover PivotFilters</span><span class="sxs-lookup"><span data-stu-id="630df-250">Remove PivotFilters</span></span>

<span data-ttu-id="630df-251">Para remover todos os PivotFilters, aplique o `clearAllFilters` método a cada PivotField, conforme mostrado no exemplo de código a seguir.</span><span class="sxs-lookup"><span data-stu-id="630df-251">To remove all PivotFilters, apply the `clearAllFilters` method to each PivotField, as shown in the following code sample.</span></span>

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

### <a name="filter-with-slicers"></a><span data-ttu-id="630df-252">Filtrar com slicers</span><span class="sxs-lookup"><span data-stu-id="630df-252">Filter with slicers</span></span>

<span data-ttu-id="630df-253">[As slicers](/javascript/api/excel/excel.slicer) permitem que os dados sejam filtrados de uma tabela Excel dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-253">[Slicers](/javascript/api/excel/excel.slicer) allow data to be filtered from an Excel PivotTable or table.</span></span> <span data-ttu-id="630df-254">Uma slicer usa valores de uma coluna especificada ou PivotField para filtrar linhas correspondentes.</span><span class="sxs-lookup"><span data-stu-id="630df-254">A slicer uses values from a specified column or PivotField to filter corresponding rows.</span></span> <span data-ttu-id="630df-255">Esses valores são armazenados [como objetos SlicerItem](/javascript/api/excel/excel.sliceritem) no `Slicer` .</span><span class="sxs-lookup"><span data-stu-id="630df-255">These values are stored as [SlicerItem](/javascript/api/excel/excel.sliceritem) objects in the `Slicer`.</span></span> <span data-ttu-id="630df-256">Seu complemento pode ajustar esses filtros, assim como os usuários ( por meio[da interface do usuário Excel interface do usuário](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span><span class="sxs-lookup"><span data-stu-id="630df-256">Your add-in can adjust these filters, as can users ([through the Excel UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span></span> <span data-ttu-id="630df-257">A slicer fica na parte superior da planilha na camada de desenho, conforme mostrado na captura de tela a seguir.</span><span class="sxs-lookup"><span data-stu-id="630df-257">The slicer sits on top of the worksheet in the drawing layer, as shown in the following screenshot.</span></span>

![Uma filtragem de dados de uma slicer em uma tabela dinâmica.](../images/excel-slicer.png)

> [!NOTE]
> <span data-ttu-id="630df-259">As técnicas descritas nesta seção se concentram em como usar slicers conectados a Tabelas Dinâmicas.</span><span class="sxs-lookup"><span data-stu-id="630df-259">The techniques described in this section focus on how to use slicers connected to PivotTables.</span></span> <span data-ttu-id="630df-260">As mesmas técnicas também se aplicam ao uso de slicers conectados a tabelas.</span><span class="sxs-lookup"><span data-stu-id="630df-260">The same techniques also apply to using slicers connected to tables.</span></span>

#### <a name="create-a-slicer"></a><span data-ttu-id="630df-261">Criar uma slicer</span><span class="sxs-lookup"><span data-stu-id="630df-261">Create a slicer</span></span>

<span data-ttu-id="630df-262">Você pode criar uma slicer em uma pasta de trabalho ou planilha usando o `Workbook.slicers.add` método ou `Worksheet.slicers.add` o método.</span><span class="sxs-lookup"><span data-stu-id="630df-262">You can create a slicer in a workbook or worksheet by using the `Workbook.slicers.add` method or `Worksheet.slicers.add` method.</span></span> <span data-ttu-id="630df-263">Isso adiciona uma slicer à [SlicerCollection](/javascript/api/excel/excel.slicercollection) do objeto `Workbook` `Worksheet` especificado ou.</span><span class="sxs-lookup"><span data-stu-id="630df-263">Doing so adds a slicer to the [SlicerCollection](/javascript/api/excel/excel.slicercollection) of the specified `Workbook` or `Worksheet` object.</span></span> <span data-ttu-id="630df-264">O `SlicerCollection.add` método tem três parâmetros:</span><span class="sxs-lookup"><span data-stu-id="630df-264">The `SlicerCollection.add` method has three parameters:</span></span>

- <span data-ttu-id="630df-265">`slicerSource`: A fonte de dados na qual a nova slicer é baseada.</span><span class="sxs-lookup"><span data-stu-id="630df-265">`slicerSource`: The data source on which the new slicer is based.</span></span> <span data-ttu-id="630df-266">Pode ser uma `PivotTable` cadeia de `Table` caracteres , ou que representa o nome ou a ID de um `PivotTable` ou `Table` .</span><span class="sxs-lookup"><span data-stu-id="630df-266">It can be a `PivotTable`, `Table`, or string representing the name or ID of a `PivotTable` or `Table`.</span></span>
- <span data-ttu-id="630df-267">`sourceField`: O campo na fonte de dados pelo qual filtrar.</span><span class="sxs-lookup"><span data-stu-id="630df-267">`sourceField`: The field in the data source by which to filter.</span></span> <span data-ttu-id="630df-268">Pode ser uma `PivotField` cadeia de `TableColumn` caracteres , ou que representa o nome ou a ID de um `PivotField` ou `TableColumn` .</span><span class="sxs-lookup"><span data-stu-id="630df-268">It can be a `PivotField`, `TableColumn`, or string representing the name or ID of a `PivotField` or `TableColumn`.</span></span>
- <span data-ttu-id="630df-269">`slicerDestination`: A planilha onde a nova slicer será criada.</span><span class="sxs-lookup"><span data-stu-id="630df-269">`slicerDestination`: The worksheet where the new slicer will be created.</span></span> <span data-ttu-id="630df-270">Pode ser um `Worksheet` objeto ou o nome ou a ID de um `Worksheet` .</span><span class="sxs-lookup"><span data-stu-id="630df-270">It can be a `Worksheet` object or the name or ID of a `Worksheet`.</span></span> <span data-ttu-id="630df-271">Esse parâmetro é desnecessário quando o `SlicerCollection` é acessado por meio de `Worksheet.slicers` .</span><span class="sxs-lookup"><span data-stu-id="630df-271">This parameter is unnecessary when the `SlicerCollection` is accessed through `Worksheet.slicers`.</span></span> <span data-ttu-id="630df-272">Nesse caso, a planilha da coleção é usada como destino.</span><span class="sxs-lookup"><span data-stu-id="630df-272">In this case, the collection's worksheet is used as the destination.</span></span>

<span data-ttu-id="630df-273">O exemplo de código a seguir adiciona uma nova slicer à **planilha** Dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-273">The following code sample adds a new slicer to the **Pivot** worksheet.</span></span> <span data-ttu-id="630df-274">A origem da slicer é a Tabela Dinâmica de Vendas do **Farm** e filtra usando os **dados Type.**</span><span class="sxs-lookup"><span data-stu-id="630df-274">The slicer's source is the **Farm Sales** PivotTable and filters using the **Type** data.</span></span> <span data-ttu-id="630df-275">A slicer também é chamada **de Fruit Slicer** para referência futura.</span><span class="sxs-lookup"><span data-stu-id="630df-275">The slicer is also named **Fruit Slicer** for future reference.</span></span>

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

#### <a name="filter-items-with-a-slicer"></a><span data-ttu-id="630df-276">Filtrar itens com uma slicer</span><span class="sxs-lookup"><span data-stu-id="630df-276">Filter items with a slicer</span></span>

<span data-ttu-id="630df-277">A slicer filtra a Tabela Dinâmica com itens do `sourceField` .</span><span class="sxs-lookup"><span data-stu-id="630df-277">The slicer filters the PivotTable with items from the `sourceField`.</span></span> <span data-ttu-id="630df-278">O `Slicer.selectItems` método define os itens que permanecem na slicer.</span><span class="sxs-lookup"><span data-stu-id="630df-278">The `Slicer.selectItems` method sets the items that remain in the slicer.</span></span> <span data-ttu-id="630df-279">Esses itens são passados para o método como `string[]` um , representando as chaves dos itens.</span><span class="sxs-lookup"><span data-stu-id="630df-279">These items are passed to the method as a `string[]`, representing the keys of the items.</span></span> <span data-ttu-id="630df-280">Todas as linhas que contêm esses itens permanecem na agregação da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-280">Any rows containing those items remain in the PivotTable's aggregation.</span></span> <span data-ttu-id="630df-281">Chamadas subsequentes `selectItems` para definir a lista como as chaves especificadas nessas chamadas.</span><span class="sxs-lookup"><span data-stu-id="630df-281">Subsequent calls to `selectItems` set the list to the keys specified in those calls.</span></span>

> [!NOTE]
> <span data-ttu-id="630df-282">Se `Slicer.selectItems` for passado um item que não está na fonte de dados, será `InvalidArgument` lançado um erro.</span><span class="sxs-lookup"><span data-stu-id="630df-282">If `Slicer.selectItems` is passed an item that's not in the data source, an `InvalidArgument` error is thrown.</span></span> <span data-ttu-id="630df-283">O conteúdo pode ser verificado por meio `Slicer.slicerItems` da propriedade, que é [uma SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span><span class="sxs-lookup"><span data-stu-id="630df-283">The contents can be verified through the `Slicer.slicerItems` property, which is a [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span></span>

<span data-ttu-id="630df-284">O exemplo de código a seguir mostra três itens que estão sendo selecionados para a slicer: **Limão,** **Limão** e **Laranja**.</span><span class="sxs-lookup"><span data-stu-id="630df-284">The following code sample shows three items being selected for the slicer: **Lemon**, **Lime**, and **Orange**.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

<span data-ttu-id="630df-285">Para remover todos os filtros da slicer, use o método, conforme `Slicer.clearFilters` mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="630df-285">To remove all filters from the slicer, use the `Slicer.clearFilters` method, as shown in the following sample.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a><span data-ttu-id="630df-286">Estilo e formatar uma slicer</span><span class="sxs-lookup"><span data-stu-id="630df-286">Style and format a slicer</span></span>

<span data-ttu-id="630df-287">Você pode ajustar as configurações de exibição de uma slicer por meio de `Slicer` propriedades.</span><span class="sxs-lookup"><span data-stu-id="630df-287">You add-in can adjust a slicer's display settings through `Slicer` properties.</span></span> <span data-ttu-id="630df-288">O exemplo de código a seguir define o estilo como **SlicerStyleLight6**, define o texto na parte superior da slicer como **Tipos** de Frutas , coloca a slicer na posição **(395, 15)** na camada de desenho e define o tamanho da slicer como **135x150** pixels.</span><span class="sxs-lookup"><span data-stu-id="630df-288">The following code sample sets the style to **SlicerStyleLight6**, sets the text at the top of the slicer to **Fruit Types**, places the slicer at the position **(395, 15)** on the drawing layer, and sets the slicer's size to **135x150** pixels.</span></span>

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

#### <a name="delete-a-slicer"></a><span data-ttu-id="630df-289">Excluir uma slicer</span><span class="sxs-lookup"><span data-stu-id="630df-289">Delete a slicer</span></span>

<span data-ttu-id="630df-290">Para excluir uma slicer, chame o `Slicer.delete` método.</span><span class="sxs-lookup"><span data-stu-id="630df-290">To delete a slicer, call the `Slicer.delete` method.</span></span> <span data-ttu-id="630df-291">O exemplo de código a seguir exclui a primeira fatia da planilha atual.</span><span class="sxs-lookup"><span data-stu-id="630df-291">The following code sample deletes the first slicer from the current worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="630df-292">Função Alterar agregação</span><span class="sxs-lookup"><span data-stu-id="630df-292">Change aggregation function</span></span>

<span data-ttu-id="630df-293">Hierarquias de dados têm seus valores agregados.</span><span class="sxs-lookup"><span data-stu-id="630df-293">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="630df-294">Para conjuntos de dados de números, essa é uma soma por padrão.</span><span class="sxs-lookup"><span data-stu-id="630df-294">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="630df-295">A `summarizeBy` propriedade define esse comportamento com base em um tipo [AggregationFunction.](/javascript/api/excel/excel.aggregationfunction)</span><span class="sxs-lookup"><span data-stu-id="630df-295">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="630df-296">Os tipos de função de agregação atualmente suportados são `Sum` , , , , , , , , , `Count` , , , `Average` , , `Max` e `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` `Automatic` (o padrão).</span><span class="sxs-lookup"><span data-stu-id="630df-296">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="630df-297">Os exemplos de código a seguir modificam a agregação como médias dos dados.</span><span class="sxs-lookup"><span data-stu-id="630df-297">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="630df-298">Alterar cálculos com um ShowAsRule</span><span class="sxs-lookup"><span data-stu-id="630df-298">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="630df-299">As Tabelas Dinâmicas, por padrão, agregam os dados de suas hierarquias de linha e coluna independentemente.</span><span class="sxs-lookup"><span data-stu-id="630df-299">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="630df-300">Um [ShowAsRule](/javascript/api/excel/excel.showasrule) altera a hierarquia de dados para valores de saída com base em outros itens na Tabela Dinâmica.</span><span class="sxs-lookup"><span data-stu-id="630df-300">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="630df-301">O `ShowAsRule` objeto tem três propriedades:</span><span class="sxs-lookup"><span data-stu-id="630df-301">The `ShowAsRule` object has three properties:</span></span>

- <span data-ttu-id="630df-302">`calculation`: O tipo de cálculo relativo a ser aplicado à hierarquia de dados (o padrão é `none` ).</span><span class="sxs-lookup"><span data-stu-id="630df-302">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
- <span data-ttu-id="630df-303">`baseField`: [PivotField](/javascript/api/excel/excel.pivotfield) na hierarquia que contém os dados base antes da aplicação do cálculo.</span><span class="sxs-lookup"><span data-stu-id="630df-303">`baseField`: The [PivotField](/javascript/api/excel/excel.pivotfield) in the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="630df-304">Como Excel tabelas dinâmicas têm um mapeamento de hierarquia para campo, você usará o mesmo nome para acessar a hierarquia e o campo.</span><span class="sxs-lookup"><span data-stu-id="630df-304">Since Excel PivotTables have a one-to-one mapping of hierarchy to field, you'll use the same name to access both the hierarchy and the field.</span></span>
- <span data-ttu-id="630df-305">`baseItem`: [PivotItem](/javascript/api/excel/excel.pivotitem) individual comparado com os valores dos campos base com base no tipo de cálculo.</span><span class="sxs-lookup"><span data-stu-id="630df-305">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="630df-306">Nem todos os cálculos exigem esse campo.</span><span class="sxs-lookup"><span data-stu-id="630df-306">Not all calculations require this field.</span></span>

<span data-ttu-id="630df-307">O exemplo a seguir define o cálculo na Soma de **Caixas Vendidas** na hierarquia de dados do Farm como uma porcentagem do total da coluna.</span><span class="sxs-lookup"><span data-stu-id="630df-307">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span>
<span data-ttu-id="630df-308">Ainda queremos que a granularidade se estenda até o nível de tipo de frutas, portanto, vamos usar a hierarquia de linha **Type** e seu campo subjacente.</span><span class="sxs-lookup"><span data-stu-id="630df-308">We still want the granularity to extend to the fruit type level, so we'll use the **Type** row hierarchy and its underlying field.</span></span>
<span data-ttu-id="630df-309">O exemplo também tem **Farm** como a hierarquia da primeira linha, portanto, o total de entradas do farm exibe a porcentagem que cada farm também é responsável por produzir.</span><span class="sxs-lookup"><span data-stu-id="630df-309">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![Uma Tabela Dinâmica mostrando as porcentagens de vendas de frutas em relação ao total geral para farms individuais e tipos de frutas individuais em cada farm.](../images/excel-pivots-showas-percentage.png)

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

<span data-ttu-id="630df-311">O exemplo anterior definiu o cálculo como a coluna, em relação ao campo de uma hierarquia de linha individual.</span><span class="sxs-lookup"><span data-stu-id="630df-311">The previous example set the calculation to the column, relative to the field of an individual row hierarchy.</span></span> <span data-ttu-id="630df-312">Quando o cálculo se refere a um item individual, use a `baseItem` propriedade.</span><span class="sxs-lookup"><span data-stu-id="630df-312">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="630df-313">O exemplo a seguir mostra o `differenceFrom` cálculo.</span><span class="sxs-lookup"><span data-stu-id="630df-313">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="630df-314">Ele exibe a diferença das entradas de hierarquia de dados de vendas de caixa de farm em relação às de **A Farms**.</span><span class="sxs-lookup"><span data-stu-id="630df-314">It displays the difference of the farm crate sales data hierarchy entries relative to those of **A Farms**.</span></span>
<span data-ttu-id="630df-315">The is Farm , so we see the differences between the other farms, well as breakdowns for each type of like fruit ( Type is also `baseField` a row hierarchy in this example). </span><span class="sxs-lookup"><span data-stu-id="630df-315">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![Uma Tabela Dinâmica mostrando as diferenças de vendas de frutas entre "A Farms" e as outras.](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="change-hierarchy-names"></a><span data-ttu-id="630df-319">Alterar nomes de hierarquia</span><span class="sxs-lookup"><span data-stu-id="630df-319">Change hierarchy names</span></span>

<span data-ttu-id="630df-320">Os campos de hierarquia são editáveis.</span><span class="sxs-lookup"><span data-stu-id="630df-320">Hierarchy fields are editable.</span></span> <span data-ttu-id="630df-321">O código a seguir demonstra como alterar os nomes exibidos de duas hierarquias de dados.</span><span class="sxs-lookup"><span data-stu-id="630df-321">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="630df-322">Confira também</span><span class="sxs-lookup"><span data-stu-id="630df-322">See also</span></span>

- [<span data-ttu-id="630df-323">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="630df-323">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="630df-324">Excel Referência da API JavaScript</span><span class="sxs-lookup"><span data-stu-id="630df-324">Excel JavaScript API Reference</span></span>](/javascript/api/excel)
