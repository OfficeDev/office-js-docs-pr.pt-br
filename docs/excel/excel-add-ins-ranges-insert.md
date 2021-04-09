---
title: Inserir intervalos usando a API JavaScript do Excel
description: Saiba como inserir um intervalo de células com a API JavaScript do Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 401a08dd10b3775012738ab9c80ec6ab367555ec
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652768"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a><span data-ttu-id="487da-103">Inserir um intervalo de células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="487da-103">Insert a range of cells using the Excel JavaScript API</span></span>

<span data-ttu-id="487da-104">Este artigo fornece um exemplo de código que insere um intervalo de células com a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="487da-104">This article provides a code sample that inserts a range of cells with the Excel JavaScript API.</span></span> <span data-ttu-id="487da-105">Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte a [classe Excel.Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="487da-105">For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="487da-106">Inserir um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="487da-106">Insert a range of cells</span></span>

<span data-ttu-id="487da-107">O exemplo de código a seguir insere um intervalo de células no local **B4:E4** e desloca outras células para baixo a fim de fornecer espaço para as novas células.</span><span class="sxs-lookup"><span data-stu-id="487da-107">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a><span data-ttu-id="487da-108">Dados antes da inserção do intervalo</span><span class="sxs-lookup"><span data-stu-id="487da-108">Data before range is inserted</span></span>

![Dados no Excel antes da inserção do intervalo](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a><span data-ttu-id="487da-110">Dados após a inserção do intervalo</span><span class="sxs-lookup"><span data-stu-id="487da-110">Data after range is inserted</span></span>

![Dados no Excel após a inserção do intervalo](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a><span data-ttu-id="487da-112">Confira também</span><span class="sxs-lookup"><span data-stu-id="487da-112">See also</span></span>

- [<span data-ttu-id="487da-113">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="487da-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="487da-114">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="487da-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="487da-115">Limpar ou excluir intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="487da-115">Clear or delete a ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-clear-delete.md)
