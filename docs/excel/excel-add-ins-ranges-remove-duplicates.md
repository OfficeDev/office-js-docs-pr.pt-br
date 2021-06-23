---
title: Remover duplicatas usando a EXCEL JavaScript
description: Saiba como usar a EXCEL JavaScript para remover duplicatas.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 859214d36bdf66a284304ba1d5f7f2d642b718cb
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075765"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a><span data-ttu-id="ae6d3-103">Remover duplicatas usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="ae6d3-103">Remove duplicates using the Excel JavaScript API</span></span>

<span data-ttu-id="ae6d3-104">Este artigo fornece um exemplo de código que remove entradas duplicadas em um intervalo usando Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-104">This article provides a code sample that removes duplicate entries in a range using the Excel JavaScript API.</span></span> <span data-ttu-id="ae6d3-105">Para ver a lista completa de propriedades e métodos que o `Range` objeto oferece suporte, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="ae6d3-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="remove-rows-with-duplicate-entries"></a><span data-ttu-id="ae6d3-106">Remover linhas com entradas duplicadas</span><span class="sxs-lookup"><span data-stu-id="ae6d3-106">Remove rows with duplicate entries</span></span>

<span data-ttu-id="ae6d3-107">O [método Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) remove linhas com entradas duplicadas nas colunas especificadas.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-107">The [Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) method removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="ae6d3-108">O método passa por cada linha no intervalo do índice de menor valor até o índice de maior valor no intervalo (de cima para baixo).</span><span class="sxs-lookup"><span data-stu-id="ae6d3-108">The method goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="ae6d3-109">Uma linha é excluída se um valor em sua coluna ou colunas especificadas aparecer mais cedo no intervalo.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-109">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="ae6d3-110">Linhas no intervalo abaixo da linha excluída são deslocadas para cima.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-110">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="ae6d3-111">`removeDuplicates` não afeta a posição de células fora do intervalo.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-111">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="ae6d3-112">`removeDuplicates` leva um `number[]` representando os índices da coluna que são verificados para duplicatas.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-112">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="ae6d3-113">Essa matriz é baseada em zero e relativa ao intervalo, não à planilha.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-113">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="ae6d3-114">O método também recebe um parâmetro booleano que especifica se a primeira linha é um header.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-114">The method also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="ae6d3-115">Quando **verdadeiro**, a primeira linha será ignorada ao considerar duplicatas.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-115">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="ae6d3-116">O método retorna um objeto que especifica o número de linhas removidas e `removeDuplicates` o número de linhas `RemoveDuplicatesResult` exclusivas restantes.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-116">The `removeDuplicates` method returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="ae6d3-117">Ao usar o método de `removeDuplicates` um intervalo, lembre-se do seguinte:</span><span class="sxs-lookup"><span data-stu-id="ae6d3-117">When using a range's `removeDuplicates` method, keep the following in mind:</span></span>

- <span data-ttu-id="ae6d3-118">`removeDuplicates` considera valores de célula, não resultados de função.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-118">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="ae6d3-119">Se as duas funções diferentes forem avaliadas como o mesmo resultado, os valores de célula não são considerados duplicatas.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-119">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="ae6d3-120">Células vazias não serão ignoradas por `removeDuplicates`.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-120">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="ae6d3-121">O valor de uma célula vazia é tratado como qualquer outro valor.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-121">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="ae6d3-122">Isso significa que as linhas vazias contidas no intervalo serão incluídas em `RemoveDuplicatesResult`.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-122">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="ae6d3-123">O exemplo de código a seguir mostra a remoção de entradas com valores duplicados na primeira coluna.</span><span class="sxs-lookup"><span data-stu-id="ae6d3-123">The following code sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

### <a name="data-before-duplicate-entries-are-removed"></a><span data-ttu-id="ae6d3-124">Dados antes que entradas duplicadas sejam removidas</span><span class="sxs-lookup"><span data-stu-id="ae6d3-124">Data before duplicate entries are removed</span></span>

![Dados em Excel antes que o método remove duplicatas do intervalo tenha sido executado.](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a><span data-ttu-id="ae6d3-126">Dados após entradas duplicadas são removidos</span><span class="sxs-lookup"><span data-stu-id="ae6d3-126">Data after duplicate entries are removed</span></span>

![Dados em Excel após a executar o método remove duplicates do intervalo.](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="ae6d3-128">Confira também</span><span class="sxs-lookup"><span data-stu-id="ae6d3-128">See also</span></span>

- [<span data-ttu-id="ae6d3-129">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ae6d3-129">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="ae6d3-130">Trabalhar com células usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="ae6d3-130">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="ae6d3-131">Intervalos de corte, cópia e colar usando a API JavaScript Excel JavaScript</span><span class="sxs-lookup"><span data-stu-id="ae6d3-131">Cut, copy, and paste ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-cut-copy-paste.md)
- [<span data-ttu-id="ae6d3-132">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="ae6d3-132">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
