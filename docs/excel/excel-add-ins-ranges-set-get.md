---
title: Definir e obter o intervalo selecionado usando a EXCEL JavaScript
description: Saiba como usar a EXCEL JavaScript para definir e obter intervalos usando a API JavaScript Excel JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0bd4a4f4bcf40e7899ee429cdc631a43ba176077
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075772"
---
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="52b2f-103">Definir e obter intervalos usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="52b2f-103">Set and get ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="52b2f-104">Este artigo fornece exemplos de código que definir e obter intervalos com a API JavaScript Excel JavaScript.</span><span class="sxs-lookup"><span data-stu-id="52b2f-104">This article provides code samples that set and get ranges with the Excel JavaScript API.</span></span> <span data-ttu-id="52b2f-105">Para ver a lista completa de propriedades e métodos que o `Range` objeto oferece suporte, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="52b2f-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="52b2f-106">Definir o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="52b2f-106">Set the selected range</span></span>

<span data-ttu-id="52b2f-107">O exemplo de código a seguir seleciona o intervalo **B2:E6** na planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="52b2f-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="52b2f-108">Intervalo selecionado B2:E6</span><span class="sxs-lookup"><span data-stu-id="52b2f-108">Selected range B2:E6</span></span>

![Intervalo selecionado Excel.](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="52b2f-110">Obter o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="52b2f-110">Get the selected range</span></span>

<span data-ttu-id="52b2f-111">O exemplo de código a seguir obtém o intervalo selecionado, carrega `address` sua propriedade e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="52b2f-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="52b2f-112">Confira também</span><span class="sxs-lookup"><span data-stu-id="52b2f-112">See also</span></span>

- [<span data-ttu-id="52b2f-113">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="52b2f-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="52b2f-114">Trabalhar com células usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="52b2f-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="52b2f-115">Definir e obter valores de intervalo, texto ou fórmulas usando Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="52b2f-115">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="52b2f-116">Definir o formato de intervalo usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="52b2f-116">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
