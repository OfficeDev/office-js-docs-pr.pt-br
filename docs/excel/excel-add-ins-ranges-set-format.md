---
title: Definir o formato de um intervalo usando a API JavaScript do Excel
description: Saiba como usar a API JavaScript do Excel para definir o formato de um intervalo.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: fdd78ea69fc38cbefb9d240dbc61554891c73c21
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652764"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a><span data-ttu-id="b5b7a-103">Definir o formato de intervalo usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="b5b7a-103">Set range format using the Excel JavaScript API</span></span>

<span data-ttu-id="b5b7a-104">Este artigo fornece exemplos de código que configuram a cor da fonte, a cor do preenchimento e o formato de número para células em um intervalo com a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="b5b7a-104">This article provides code samples that set font color, fill color, and number format for cells in a range with the Excel JavaScript API.</span></span> <span data-ttu-id="b5b7a-105">Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="b5b7a-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a><span data-ttu-id="b5b7a-106">Definir cor da fonte e cor de preenchimento</span><span class="sxs-lookup"><span data-stu-id="b5b7a-106">Set font color and fill color</span></span>

<span data-ttu-id="b5b7a-107">O exemplo de código a seguir define a cor da fonte e a cor de preenchimento para células no intervalo **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="b5b7a-107">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a><span data-ttu-id="b5b7a-108">Dados no intervalo antes da definição da cor da fonte e da cor de preenchimento</span><span class="sxs-lookup"><span data-stu-id="b5b7a-108">Data in range before font color and fill color are set</span></span>

![Dados no Excel antes da definição do formato](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a><span data-ttu-id="b5b7a-110">Dados no intervalo após a definição da cor da fonte e da cor de preenchimento</span><span class="sxs-lookup"><span data-stu-id="b5b7a-110">Data in range after font color and fill color are set</span></span>

![Dados no Excel após a definição do formato](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a><span data-ttu-id="b5b7a-112">Definir formato de número</span><span class="sxs-lookup"><span data-stu-id="b5b7a-112">Set number format</span></span>

<span data-ttu-id="b5b7a-113">O exemplo de código a seguir define o formato de número para as células no intervalo **D3:E5**.</span><span class="sxs-lookup"><span data-stu-id="b5b7a-113">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    var range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-number-format-is-set"></a><span data-ttu-id="b5b7a-114">Dados no intervalo antes da definição do formato de número</span><span class="sxs-lookup"><span data-stu-id="b5b7a-114">Data in range before number format is set</span></span>

![Dados no Excel antes que o formato de número seja definido](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a><span data-ttu-id="b5b7a-116">Dados no intervalo após a definição do formato de número</span><span class="sxs-lookup"><span data-stu-id="b5b7a-116">Data in range after number format is set</span></span>

![Dados no Excel após o formato de número ser definido](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a><span data-ttu-id="b5b7a-118">Confira também</span><span class="sxs-lookup"><span data-stu-id="b5b7a-118">See also</span></span>

- [<span data-ttu-id="b5b7a-119">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b5b7a-119">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="b5b7a-120">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="b5b7a-120">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="b5b7a-121">Definir e obter intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="b5b7a-121">Set and get ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get.md)
- [<span data-ttu-id="b5b7a-122">Definir e obter valores de intervalo, texto ou fórmulas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="b5b7a-122">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
