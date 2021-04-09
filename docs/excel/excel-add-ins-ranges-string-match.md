---
title: Encontre uma cadeia de caracteres usando a API JavaScript do Excel
description: Saiba como encontrar uma cadeia de caracteres em um intervalo usando a API JavaScript do Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9b649bb249cd24d7578bc4f8285e5d0a23d0e4cd
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652757"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="48631-103">Encontre uma cadeia de caracteres em um intervalo usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="48631-103">Find a string within a range using the Excel JavaScript API</span></span>

<span data-ttu-id="48631-104">Este artigo fornece um exemplo de código que localiza uma cadeia de caracteres dentro de um intervalo usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="48631-104">This article provides a code sample that finds a string within a range using the Excel JavaScript API.</span></span> <span data-ttu-id="48631-105">Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="48631-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a><span data-ttu-id="48631-106">Corresponder a uma cadeia de caracteres dentro de um intervalo</span><span class="sxs-lookup"><span data-stu-id="48631-106">Match a string within a range</span></span>

<span data-ttu-id="48631-107">O objeto `Range` tem um método `find` para pesquisar uma cadeia especificada dentro do intervalo.</span><span class="sxs-lookup"><span data-stu-id="48631-107">The `Range` object has a `find` method to search for a specified string within the range.</span></span> <span data-ttu-id="48631-108">Ele retorna o intervalo da primeira célula com o texto correspondente.</span><span class="sxs-lookup"><span data-stu-id="48631-108">It returns the range of the first cell with matching text.</span></span>

<span data-ttu-id="48631-109">O exemplo de código a seguir localiza a primeira célula com um valor igual à cadeia de caracteres **Alimentos** e registra o seu endereço no console.</span><span class="sxs-lookup"><span data-stu-id="48631-109">The following code sample finds the first cell with a value equal to the string **Food** and logs its address to the console.</span></span> <span data-ttu-id="48631-110">Observe que `find` exibe um erro `ItemNotFound` se a cadeia de caracteres especificada não existir no intervalo.</span><span class="sxs-lookup"><span data-stu-id="48631-110">Note that `find` throws an `ItemNotFound` error if the specified string doesn't exist in the range.</span></span> <span data-ttu-id="48631-111">Se você acha que a cadeia de caracteres especificada pode não estar no intervalo, use o método [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) para que seu código manipule normalmente esse cenário.</span><span class="sxs-lookup"><span data-stu-id="48631-111">If you expect that the specified string may not exist in the range, use the [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="48631-112">Quando o método `find` é chamado em um intervalo que representa uma única célula, a planilha inteira é pesquisada.</span><span class="sxs-lookup"><span data-stu-id="48631-112">When the `find` method is called on a range representing a single cell, the entire worksheet is searched.</span></span> <span data-ttu-id="48631-113">A pesquisa começa na célula e segue na direção especificada pelo `SearchCriteria.searchDirection`, envolvendo as extremidades da planilha, se necessário.</span><span class="sxs-lookup"><span data-stu-id="48631-113">The search begins at that cell and goes in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.</span></span>

## <a name="see-also"></a><span data-ttu-id="48631-114">Confira também</span><span class="sxs-lookup"><span data-stu-id="48631-114">See also</span></span>

- [<span data-ttu-id="48631-115">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="48631-115">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="48631-116">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="48631-116">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="48631-117">Encontre células especiais em um intervalo usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="48631-117">Find special cells within a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-special-cells.md)
