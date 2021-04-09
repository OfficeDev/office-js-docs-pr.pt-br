---
title: Definir e obter o intervalo selecionado usando a API JavaScript do Excel
description: Saiba como usar a API JavaScript do Excel para definir e obter intervalos usando a API JavaScript do Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 06b6219924f0667ecef57d608cb417a76ef8031d
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652755"
---
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="643a1-103">Definir e obter intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="643a1-103">Set and get ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="643a1-104">Este artigo fornece exemplos de código que definir e obter intervalos com a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="643a1-104">This article provides code samples that set and get ranges with the Excel JavaScript API.</span></span> <span data-ttu-id="643a1-105">Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="643a1-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="643a1-106">Definir o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="643a1-106">Set the selected range</span></span>

<span data-ttu-id="643a1-107">O exemplo de código a seguir seleciona o intervalo **B2:E6** na planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="643a1-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="643a1-108">Intervalo selecionado B2:E6</span><span class="sxs-lookup"><span data-stu-id="643a1-108">Selected range B2:E6</span></span>

![Intervalo selecionado no Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="643a1-110">Obter o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="643a1-110">Get the selected range</span></span>

<span data-ttu-id="643a1-111">O exemplo de código a seguir obtém o intervalo selecionado, carrega `address` sua propriedade e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="643a1-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="643a1-112">Confira também</span><span class="sxs-lookup"><span data-stu-id="643a1-112">See also</span></span>

- [<span data-ttu-id="643a1-113">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="643a1-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="643a1-114">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="643a1-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="643a1-115">Definir e obter valores de intervalo, texto ou fórmulas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="643a1-115">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="643a1-116">Definir o formato de intervalo usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="643a1-116">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
