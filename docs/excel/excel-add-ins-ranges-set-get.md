---
title: Definir e obter o intervalo selecionado usando a EXCEL JavaScript
description: Saiba como usar a EXCEL JavaScript para definir e obter o intervalo selecionado usando Excel API JavaScript.
ms.date: 07/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 623ba5c1b9e76151d4a2c4b169e655236b37e8c8
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290779"
---
# <a name="set-and-get-the-selected-range-using-the-excel-javascript-api"></a><span data-ttu-id="29c55-103">Definir e obter o intervalo selecionado usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="29c55-103">Set and get the selected range using the Excel JavaScript API</span></span>

<span data-ttu-id="29c55-104">Este artigo fornece exemplos de código que configuram e selecionam o intervalo com a API JavaScript Excel javascript.</span><span class="sxs-lookup"><span data-stu-id="29c55-104">This article provides code samples that set and get the selected range with the Excel JavaScript API.</span></span> <span data-ttu-id="29c55-105">Para ver a lista completa de propriedades e métodos que o `Range` objeto oferece suporte, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="29c55-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="29c55-106">Definir o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="29c55-106">Set the selected range</span></span>

<span data-ttu-id="29c55-107">O exemplo de código a seguir seleciona o intervalo **B2:E6** na planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="29c55-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="29c55-108">Intervalo selecionado B2:E6</span><span class="sxs-lookup"><span data-stu-id="29c55-108">Selected range B2:E6</span></span>

![Intervalo selecionado Excel.](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="29c55-110">Obter o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="29c55-110">Get the selected range</span></span>

<span data-ttu-id="29c55-111">O exemplo de código a seguir obtém o intervalo selecionado, carrega `address` sua propriedade e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="29c55-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="select-the-edge-of-a-used-range"></a><span data-ttu-id="29c55-112">Selecione a borda de um intervalo usado</span><span class="sxs-lookup"><span data-stu-id="29c55-112">Select the edge of a used range</span></span>

<span data-ttu-id="29c55-113">Os [métodos Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) e [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) permitem que o seu complemento replique o comportamento dos atalhos de seleção do teclado, selecionando a borda do intervalo usado com base no intervalo selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="29c55-113">The [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) and [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) methods let your add-in replicate the behavior of the keyboard selection shortcuts, selecting the edge of the used range based on the currently selected range.</span></span> <span data-ttu-id="29c55-114">Para saber mais sobre intervalos usados, consulte [Obter intervalo usado](excel-add-ins-ranges-get.md#get-used-range).</span><span class="sxs-lookup"><span data-stu-id="29c55-114">To learn more about used ranges, see [Get used range](excel-add-ins-ranges-get.md#get-used-range).</span></span>

<span data-ttu-id="29c55-115">Na captura de tela a seguir, o intervalo usado é a tabela com valores em cada célula, **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="29c55-115">In the following screenshot, the used range is the table with values in each cell, **C5:F12**.</span></span> <span data-ttu-id="29c55-116">As células vazias fora desta tabela estão fora do intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="29c55-116">The empty cells outside this table are outside the used range.</span></span>

![Uma tabela com dados de C5:F12 em Excel.](../images/excel-ranges-used-range.png)

### <a name="select-the-cell-at-the-edge-of-the-current-used-range"></a><span data-ttu-id="29c55-118">Selecione a célula na borda do intervalo usado atual</span><span class="sxs-lookup"><span data-stu-id="29c55-118">Select the cell at the edge of the current used range</span></span>

<span data-ttu-id="29c55-119">O exemplo de código a seguir mostra como usar o método para selecionar a célula na borda mais distante do intervalo usado `Range.getRangeEdge` atual, na direção para cima.</span><span class="sxs-lookup"><span data-stu-id="29c55-119">The following code sample shows how use the `Range.getRangeEdge` method to select the cell at the furthest edge of the current used range, in the direction up.</span></span> <span data-ttu-id="29c55-120">Essa ação corresponde ao resultado do uso do atalho do teclado de tecla de seta Ctrl+Up enquanto um intervalo é selecionado.</span><span class="sxs-lookup"><span data-stu-id="29c55-120">This action matches the result of using the Ctrl+Up arrow key keyboard shortcut while a range is selected.</span></span>

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

#### <a name="before-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="29c55-121">Antes de selecionar a célula na borda do intervalo usado</span><span class="sxs-lookup"><span data-stu-id="29c55-121">Before selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="29c55-122">A captura de tela a seguir mostra um intervalo usado e um intervalo selecionado dentro do intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="29c55-122">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="29c55-123">O intervalo usado é uma tabela com dados **em C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="29c55-123">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="29c55-124">Dentro desta tabela, o intervalo **D8:E9** está selecionado.</span><span class="sxs-lookup"><span data-stu-id="29c55-124">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="29c55-125">Essa seleção é o *estado anterior,* antes de executar o `Range.getRangeEdge` método.</span><span class="sxs-lookup"><span data-stu-id="29c55-125">This selection is the *before* state, prior to running the `Range.getRangeEdge` method.</span></span>

![Uma tabela com dados de C5:F12 em Excel.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="29c55-128">Depois de selecionar a célula na borda do intervalo usado</span><span class="sxs-lookup"><span data-stu-id="29c55-128">After selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="29c55-129">A captura de tela a seguir mostra a mesma tabela da captura de tela anterior, com dados no intervalo **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="29c55-129">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="29c55-130">Dentro desta tabela, o intervalo **D5** é selecionado.</span><span class="sxs-lookup"><span data-stu-id="29c55-130">Inside this table, the range **D5** is selected.</span></span> <span data-ttu-id="29c55-131">Essa seleção é *após o* estado, depois de executar o método para selecionar a célula na borda do intervalo usado na direção `Range.getRangeEdge` para cima.</span><span class="sxs-lookup"><span data-stu-id="29c55-131">This selection is *after* state, after running the `Range.getRangeEdge` method to select the cell at the edge of the used range in the up direction.</span></span>

![Uma tabela com dados de C5:F12 em Excel.](../images/excel-ranges-used-range-d5.png)

### <a name="select-all-cells-from-current-range-to-furthest-edge-of-used-range"></a><span data-ttu-id="29c55-134">Selecione todas as células do intervalo atual até a borda mais distante do intervalo usado</span><span class="sxs-lookup"><span data-stu-id="29c55-134">Select all cells from current range to furthest edge of used range</span></span>

<span data-ttu-id="29c55-135">O exemplo de código a seguir mostra como usar o método para selecionar todas as células do intervalo selecionado no momento até a borda mais distante do intervalo usado, na direção `Range.getExtendedRange` para baixo.</span><span class="sxs-lookup"><span data-stu-id="29c55-135">The following code sample shows how use the `Range.getExtendedRange` method to to select all the cells from the currently selected range to the furthest edge of the used range, in the direction down.</span></span> <span data-ttu-id="29c55-136">Essa ação corresponde ao resultado do uso do atalho do teclado de tecla de seta Ctrl+Shift+Down enquanto um intervalo é selecionado.</span><span class="sxs-lookup"><span data-stu-id="29c55-136">This action matches the result of using the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.</span></span>

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

#### <a name="before-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="29c55-137">Antes de selecionar todas as células do intervalo atual até a borda do intervalo usado</span><span class="sxs-lookup"><span data-stu-id="29c55-137">Before selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="29c55-138">A captura de tela a seguir mostra um intervalo usado e um intervalo selecionado dentro do intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="29c55-138">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="29c55-139">O intervalo usado é uma tabela com dados **em C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="29c55-139">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="29c55-140">Dentro desta tabela, o intervalo **D8:E9** está selecionado.</span><span class="sxs-lookup"><span data-stu-id="29c55-140">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="29c55-141">Essa seleção é o *estado anterior,* antes de executar o `Range.getExtendedRange` método.</span><span class="sxs-lookup"><span data-stu-id="29c55-141">This selection is the *before* state, prior to running the `Range.getExtendedRange` method.</span></span>

![Uma tabela com dados de C5:F12 em Excel.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="29c55-144">Depois de selecionar todas as células do intervalo atual até a borda do intervalo usado</span><span class="sxs-lookup"><span data-stu-id="29c55-144">After selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="29c55-145">A captura de tela a seguir mostra a mesma tabela da captura de tela anterior, com dados no intervalo **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="29c55-145">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="29c55-146">Dentro desta tabela, o intervalo **D8:E12** está selecionado.</span><span class="sxs-lookup"><span data-stu-id="29c55-146">Inside this table, the range **D8:E12** is selected.</span></span> <span data-ttu-id="29c55-147">Essa seleção é *após o* estado, depois de executar o método para selecionar todas as células do intervalo atual até a borda do intervalo usado na direção `Range.getExtendedRange` para baixo.</span><span class="sxs-lookup"><span data-stu-id="29c55-147">This selection is *after* state, after running the `Range.getExtendedRange` method to select all the cells from the current range to the edge of the used range in the down direction.</span></span>

![Uma tabela com dados de C5:F12 em Excel.](../images/excel-ranges-used-range-d8-e12.png)

## <a name="see-also"></a><span data-ttu-id="29c55-150">Confira também</span><span class="sxs-lookup"><span data-stu-id="29c55-150">See also</span></span>

- [<span data-ttu-id="29c55-151">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="29c55-151">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="29c55-152">Trabalhar com células usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="29c55-152">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="29c55-153">Definir e obter valores de intervalo, texto ou fórmulas usando Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="29c55-153">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="29c55-154">Definir o formato de intervalo usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="29c55-154">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
