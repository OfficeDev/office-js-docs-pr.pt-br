---
title: Definir e obter o intervalo selecionado usando a EXCEL JavaScript
description: Saiba como usar a EXCEL JavaScript para definir e obter o intervalo selecionado usando Excel API JavaScript.
ms.date: 06/22/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9e4c31f165b39d45fac342cb85577ef737105472
ms.sourcegitcommit: ebb4a22a0bdeb5623c72b9494ebbce3909d0c90c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2021
ms.locfileid: "53126717"
---
# <a name="set-and-get-the-selected-range-using-the-excel-javascript-api"></a><span data-ttu-id="5c5de-103">Definir e obter o intervalo selecionado usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="5c5de-103">Set and get the selected range using the Excel JavaScript API</span></span>

<span data-ttu-id="5c5de-104">Este artigo fornece exemplos de código que configuram e selecionam o intervalo com a API JavaScript Excel javascript.</span><span class="sxs-lookup"><span data-stu-id="5c5de-104">This article provides code samples that set and get the selected range with the Excel JavaScript API.</span></span> <span data-ttu-id="5c5de-105">Para ver a lista completa de propriedades e métodos que o `Range` objeto oferece suporte, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="5c5de-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="5c5de-106">Definir o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="5c5de-106">Set the selected range</span></span>

<span data-ttu-id="5c5de-107">O exemplo de código a seguir seleciona o intervalo **B2:E6** na planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="5c5de-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="5c5de-108">Intervalo selecionado B2:E6</span><span class="sxs-lookup"><span data-stu-id="5c5de-108">Selected range B2:E6</span></span>

![Intervalo selecionado Excel.](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="5c5de-110">Obter o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="5c5de-110">Get the selected range</span></span>

<span data-ttu-id="5c5de-111">O exemplo de código a seguir obtém o intervalo selecionado, carrega `address` sua propriedade e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="5c5de-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="select-the-edge-of-a-used-range-online-only"></a><span data-ttu-id="5c5de-112">Selecione a borda de um intervalo usado (somente online)</span><span class="sxs-lookup"><span data-stu-id="5c5de-112">Select the edge of a used range (online-only)</span></span>

> [!NOTE]
> <span data-ttu-id="5c5de-113">No momento, os métodos e estão disponíveis `Range.getRangeEdge` `Range.getExtendedRange` apenas no ExcelApiOnline 1.1.</span><span class="sxs-lookup"><span data-stu-id="5c5de-113">The `Range.getRangeEdge` and `Range.getExtendedRange` methods are currently only available in ExcelApiOnline 1.1.</span></span> <span data-ttu-id="5c5de-114">Para saber mais, consulte Excel conjunto de [requisitos somente da API JavaScript online.](../reference/requirement-sets/excel-api-online-requirement-set.md)</span><span class="sxs-lookup"><span data-stu-id="5c5de-114">To learn more, see [Excel JavaScript API online-only requirement set](../reference/requirement-sets/excel-api-online-requirement-set.md).</span></span>

<span data-ttu-id="5c5de-115">Os [métodos Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) e [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) permitem que o seu complemento replique o comportamento dos atalhos de seleção do teclado, selecionando a borda do intervalo usado com base no intervalo selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="5c5de-115">The [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) and [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) methods let your add-in replicate the behavior of the keyboard selection shortcuts, selecting the edge of the used range based on the currently selected range.</span></span> <span data-ttu-id="5c5de-116">Para saber mais sobre intervalos usados, consulte [Obter intervalo usado](excel-add-ins-ranges-get.md#get-used-range).</span><span class="sxs-lookup"><span data-stu-id="5c5de-116">To learn more about used ranges, see [Get used range](excel-add-ins-ranges-get.md#get-used-range).</span></span>

<span data-ttu-id="5c5de-117">Na captura de tela a seguir, o intervalo usado é a tabela com valores em cada célula, **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="5c5de-117">In the following screenshot, the used range is the table with values in each cell, **C5:F12**.</span></span> <span data-ttu-id="5c5de-118">As células vazias fora desta tabela estão fora do intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="5c5de-118">The empty cells outside this table are outside the used range.</span></span>

![Uma tabela com dados de C5:F12 em Excel.](../images/excel-ranges-used-range.png)

### <a name="select-the-cell-at-the-edge-of-the-current-used-range"></a><span data-ttu-id="5c5de-120">Selecione a célula na borda do intervalo usado atual</span><span class="sxs-lookup"><span data-stu-id="5c5de-120">Select the cell at the edge of the current used range</span></span>

<span data-ttu-id="5c5de-121">O exemplo de código a seguir mostra como usar o método para selecionar a célula na borda mais distante do intervalo usado `Range.getRangeEdge` atual, na direção para cima.</span><span class="sxs-lookup"><span data-stu-id="5c5de-121">The following code sample shows how use the `Range.getRangeEdge` method to select the cell at the furthest edge of the current used range, in the direction up.</span></span> <span data-ttu-id="5c5de-122">Essa ação corresponde ao resultado do uso do atalho do teclado de tecla de seta Ctrl+Up enquanto um intervalo é selecionado.</span><span class="sxs-lookup"><span data-stu-id="5c5de-122">This action matches the result of using the Ctrl+Up arrow key keyboard shortcut while a range is selected.</span></span>

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.up;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get the top-most cell of the current used range.
    // This method acts like the Ctrl+Up arrow key keyboard shortcut while a range is selected.
    var rangeEdge = range.getRangeEdge(
      direction,
      activeCell
    );
    rangeEdge.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="5c5de-123">Antes de selecionar a célula na borda do intervalo usado</span><span class="sxs-lookup"><span data-stu-id="5c5de-123">Before selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="5c5de-124">A captura de tela a seguir mostra um intervalo usado e um intervalo selecionado dentro do intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="5c5de-124">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="5c5de-125">O intervalo usado é uma tabela com dados **em C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="5c5de-125">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="5c5de-126">Dentro desta tabela, o intervalo **D8:E9** está selecionado.</span><span class="sxs-lookup"><span data-stu-id="5c5de-126">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="5c5de-127">Essa seleção é o *estado anterior,* antes de executar o `Range.getRangeEdge` método.</span><span class="sxs-lookup"><span data-stu-id="5c5de-127">This selection is the *before* state, prior to running the `Range.getRangeEdge` method.</span></span>

![Uma tabela com dados de C5:F12 em Excel.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="5c5de-130">Depois de selecionar a célula na borda do intervalo usado</span><span class="sxs-lookup"><span data-stu-id="5c5de-130">After selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="5c5de-131">A captura de tela a seguir mostra a mesma tabela da captura de tela anterior, com dados no intervalo **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="5c5de-131">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="5c5de-132">Dentro desta tabela, o intervalo **D5** é selecionado.</span><span class="sxs-lookup"><span data-stu-id="5c5de-132">Inside this table, the range **D5** is selected.</span></span> <span data-ttu-id="5c5de-133">Essa seleção é *após o* estado, depois de executar o método para selecionar a célula na borda do intervalo usado na direção `Range.getRangeEdge` para cima.</span><span class="sxs-lookup"><span data-stu-id="5c5de-133">This selection is *after* state, after running the `Range.getRangeEdge` method to select the cell at the edge of the used range in the up direction.</span></span>

![Uma tabela com dados de C5:F12 em Excel.](../images/excel-ranges-used-range-d5.png)

### <a name="select-all-cells-from-current-range-to-furthest-edge-of-used-range"></a><span data-ttu-id="5c5de-136">Selecione todas as células do intervalo atual até a borda mais distante do intervalo usado</span><span class="sxs-lookup"><span data-stu-id="5c5de-136">Select all cells from current range to furthest edge of used range</span></span>

<span data-ttu-id="5c5de-137">O exemplo de código a seguir mostra como usar o método para selecionar todas as células do intervalo selecionado no momento até a borda mais distante do intervalo usado, na direção `Range.getExtendedRange` para baixo.</span><span class="sxs-lookup"><span data-stu-id="5c5de-137">The following code sample shows how use the `Range.getExtendedRange` method to to select all the cells from the currently selected range to the furthest edge of the used range, in the direction down.</span></span> <span data-ttu-id="5c5de-138">Essa ação corresponde ao resultado do uso do atalho do teclado de tecla de seta Ctrl+Shift+Down enquanto um intervalo é selecionado.</span><span class="sxs-lookup"><span data-stu-id="5c5de-138">This action matches the result of using the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.</span></span>

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.down;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get all the cells from the currently selected range to the bottom-most edge of the used range.
    // This method acts like the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.
    var extendedRange = range.getExtendedRange(
      direction,
      activeCell
    );
    extendedRange.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="5c5de-139">Antes de selecionar todas as células do intervalo atual até a borda do intervalo usado</span><span class="sxs-lookup"><span data-stu-id="5c5de-139">Before selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="5c5de-140">A captura de tela a seguir mostra um intervalo usado e um intervalo selecionado dentro do intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="5c5de-140">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="5c5de-141">O intervalo usado é uma tabela com dados **em C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="5c5de-141">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="5c5de-142">Dentro desta tabela, o intervalo **D8:E9** está selecionado.</span><span class="sxs-lookup"><span data-stu-id="5c5de-142">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="5c5de-143">Essa seleção é o *estado anterior,* antes de executar o `Range.getExtendedRange` método.</span><span class="sxs-lookup"><span data-stu-id="5c5de-143">This selection is the *before* state, prior to running the `Range.getExtendedRange` method.</span></span>

![Uma tabela com dados de C5:F12 em Excel.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="5c5de-146">Depois de selecionar todas as células do intervalo atual até a borda do intervalo usado</span><span class="sxs-lookup"><span data-stu-id="5c5de-146">After selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="5c5de-147">A captura de tela a seguir mostra a mesma tabela da captura de tela anterior, com dados no intervalo **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="5c5de-147">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="5c5de-148">Dentro desta tabela, o intervalo **D8:E12** está selecionado.</span><span class="sxs-lookup"><span data-stu-id="5c5de-148">Inside this table, the range **D8:E12** is selected.</span></span> <span data-ttu-id="5c5de-149">Essa seleção é *após o* estado, depois de executar o método para selecionar todas as células do intervalo atual até a borda do intervalo usado na direção `Range.getExtendedRange` para baixo.</span><span class="sxs-lookup"><span data-stu-id="5c5de-149">This selection is *after* state, after running the `Range.getExtendedRange` method to select all the cells from the current range to the edge of the used range in the down direction.</span></span>

![Uma tabela com dados de C5:F12 em Excel.](../images/excel-ranges-used-range-d8-e12.png)

## <a name="see-also"></a><span data-ttu-id="5c5de-152">Confira também</span><span class="sxs-lookup"><span data-stu-id="5c5de-152">See also</span></span>

- [<span data-ttu-id="5c5de-153">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5c5de-153">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="5c5de-154">Trabalhar com células usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="5c5de-154">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="5c5de-155">Definir e obter valores de intervalo, texto ou fórmulas usando Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="5c5de-155">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="5c5de-156">Definir o formato de intervalo usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="5c5de-156">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
