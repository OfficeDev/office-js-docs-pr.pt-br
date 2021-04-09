---
title: Limpar ou excluir intervalos usando a API JavaScript do Excel
description: Saiba como limpar ou excluir intervalos usando a API JavaScript do Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7e030c6b5ba7ba6e6c54e9be0524cd93c2516bcb
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652783"
---
# <a name="clear-or-delete-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="be238-103">Limpar ou excluir intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="be238-103">Clear or delete ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="be238-104">Este artigo fornece exemplos de código que limpam e excluem intervalos com a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="be238-104">This article provides code samples that clear and delete ranges with the Excel JavaScript API.</span></span> <span data-ttu-id="be238-105">Para ver a lista completa de propriedades e métodos suportados pelo `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="be238-105">For the complete list of properties and methods supported by the `Range` object, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="be238-106">Limpar um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="be238-106">Clear a range of cells</span></span>

<span data-ttu-id="be238-107">O exemplo de código a seguir limpa todo o conteúdo e a formatação das células no intervalo **E2:E5**.</span><span class="sxs-lookup"><span data-stu-id="be238-107">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-cleared"></a><span data-ttu-id="be238-108">Dados antes da limpeza do intervalo</span><span class="sxs-lookup"><span data-stu-id="be238-108">Data before range is cleared</span></span>

![Dados no Excel antes da limpeza do intervalo](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a><span data-ttu-id="be238-110">Dados após a limpeza do intervalo</span><span class="sxs-lookup"><span data-stu-id="be238-110">Data after range is cleared</span></span>

![Dados no Excel após a limpeza do intervalo](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="be238-112">Excluir um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="be238-112">Delete a range of cells</span></span>

<span data-ttu-id="be238-113">O exemplo de código a seguir exclui as células no intervalo **B4:E4** e desloca outras células para cima para preencher o espaço que foi desocupado pelas células excluídas.</span><span class="sxs-lookup"><span data-stu-id="be238-113">The following code sample deletes the cells in the range **B4:E4** and shifts other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-deleted"></a><span data-ttu-id="be238-114">Dados antes da exclusão do intervalo</span><span class="sxs-lookup"><span data-stu-id="be238-114">Data before range is deleted</span></span>

![Dados no Excel antes da exclusão do intervalo](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a><span data-ttu-id="be238-116">Dados após a exclusão do intervalo</span><span class="sxs-lookup"><span data-stu-id="be238-116">Data after range is deleted</span></span>

![Dados no Excel após a exclusão do intervalo](../images/excel-ranges-after-delete.png)


## <a name="see-also"></a><span data-ttu-id="be238-118">Confira também</span><span class="sxs-lookup"><span data-stu-id="be238-118">See also</span></span>

- [<span data-ttu-id="be238-119">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="be238-119">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="be238-120">Definir e obter intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="be238-120">Set and get ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get.md)
- [<span data-ttu-id="be238-121">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="be238-121">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
