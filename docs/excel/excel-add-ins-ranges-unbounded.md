---
title: Ler ou gravar em um intervalo não-rebote usando a API JavaScript do Excel
description: Saiba como usar a API JavaScript do Excel para ler ou gravar em um intervalo não-rebote.
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f7be2efc3e069ea3451088608ca5255a632ef863
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652756"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a><span data-ttu-id="505b4-103">Ler ou gravar em um intervalo não-rebote usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="505b4-103">Read or write to an unbounded range using the Excel JavaScript API</span></span>

<span data-ttu-id="505b4-104">Este artigo descreve como ler e gravar em um intervalo não-rebote com a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="505b4-104">This article describes how to read and write to an unbounded range with the Excel JavaScript API.</span></span> <span data-ttu-id="505b4-105">Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="505b4-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

<span data-ttu-id="505b4-106">Um endereço de intervalo não rebotado é um endereço de intervalo que especifica colunas inteiras ou linhas inteiras.</span><span class="sxs-lookup"><span data-stu-id="505b4-106">An unbounded range address is a range address that specifies either entire columns or entire rows.</span></span> <span data-ttu-id="505b4-107">Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="505b4-107">For example:</span></span>

- <span data-ttu-id="505b4-108">Endereços de intervalo compostos por colunas inteiras:</span><span class="sxs-lookup"><span data-stu-id="505b4-108">Range addresses comprised of entire columns:</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
- <span data-ttu-id="505b4-109">Endereços de intervalo compostos por linhas inteiras:</span><span class="sxs-lookup"><span data-stu-id="505b4-109">Range addresses comprised of entire rows:</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>

## <a name="read-an-unbounded-range"></a><span data-ttu-id="505b4-110">Ler um intervalo não limitado</span><span class="sxs-lookup"><span data-stu-id="505b4-110">Read an unbounded range</span></span>

<span data-ttu-id="505b4-p103">Quando uma API faz uma solicitação para recuperar um intervalo não limitado (por exemplo, `getRange('C:C')`), a resposta conterá valores `null` para as propriedades no nível de célula, como `values`, `text`, `numberFormat` e `formula`. Outras propriedades do intervalo, como `address` e `cellCount`, conterão valores válidos para o intervalo não limitado.</span><span class="sxs-lookup"><span data-stu-id="505b4-p103">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>

## <a name="write-to-an-unbounded-range"></a><span data-ttu-id="505b4-113">Gravar em um intervalo não limitado</span><span class="sxs-lookup"><span data-stu-id="505b4-113">Write to an unbounded range</span></span>

<span data-ttu-id="505b4-114">Não é possível definir propriedades no nível da célula, como , e em um intervalo não rebotado porque a solicitação de `values` `numberFormat` entrada é muito `formula` grande.</span><span class="sxs-lookup"><span data-stu-id="505b4-114">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on an unbounded range because the input request is too large.</span></span> <span data-ttu-id="505b4-115">Por exemplo, o exemplo de código a seguir não é válido porque ele tenta especificar para um `values` intervalo não-rebote.</span><span class="sxs-lookup"><span data-stu-id="505b4-115">For example, the following code example is not valid because it attempts to specify `values` for an unbounded range.</span></span> <span data-ttu-id="505b4-116">A API retornará um erro se você tentar definir propriedades no nível da célula para um intervalo não-rebote.</span><span class="sxs-lookup"><span data-stu-id="505b4-116">The API returns an error if you attempt to set cell-level properties for an unbounded range.</span></span>

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a><span data-ttu-id="505b4-117">Confira também</span><span class="sxs-lookup"><span data-stu-id="505b4-117">See also</span></span>

- [<span data-ttu-id="505b4-118">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="505b4-118">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="505b4-119">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="505b4-119">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="505b4-120">Ler ou gravar em um intervalo grande usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="505b4-120">Read or write to a large range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-large.md)
- [<span data-ttu-id="505b4-121">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="505b4-121">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
