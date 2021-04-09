---
title: Intervalos de grupo usando a API JavaScript do Excel
description: Saiba como agrupar linhas ou colunas de um intervalo para criar um contorno usando a API JavaScript do Excel.
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 32f65cf88c23bd6368b37318d3ba20fde95b8436
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652769"
---
# <a name="group-ranges-for-an-outline-using-the-excel-javascript-api"></a><span data-ttu-id="01949-103">Intervalos de grupo para um contorno usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="01949-103">Group ranges for an outline using the Excel JavaScript API</span></span>

<span data-ttu-id="01949-104">Este artigo fornece um exemplo de código que mostra como agrupar intervalos para um contorno usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="01949-104">This article provides a code sample that shows how to group ranges for an outline using the Excel JavaScript API.</span></span> <span data-ttu-id="01949-105">Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="01949-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="group-rows-or-columns-of-a-range-for-an-outline"></a><span data-ttu-id="01949-106">Linhas de grupo ou colunas de um intervalo para um contorno</span><span class="sxs-lookup"><span data-stu-id="01949-106">Group rows or columns of a range for an outline</span></span>

<span data-ttu-id="01949-107">Linhas ou colunas de um intervalo podem ser agrupadas para criar um [contorno](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF).</span><span class="sxs-lookup"><span data-stu-id="01949-107">Rows or columns of a range can be grouped together to create an [outline](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF).</span></span> <span data-ttu-id="01949-108">Esses grupos podem ser recolhidos e expandidos para ocultar e mostrar as células correspondentes.</span><span class="sxs-lookup"><span data-stu-id="01949-108">These groups can be collapsed and expanded to hide and show the corresponding cells.</span></span> <span data-ttu-id="01949-109">Isso facilita a análise rápida dos dados de linha superior.</span><span class="sxs-lookup"><span data-stu-id="01949-109">This makes quick analysis of top-line data easier.</span></span> <span data-ttu-id="01949-110">Use [Range.group](/javascript/api/excel/excel.range#group-groupoption-) para fazer esses grupos de contornos.</span><span class="sxs-lookup"><span data-stu-id="01949-110">Use [Range.group](/javascript/api/excel/excel.range#group-groupoption-) to make these outline groups.</span></span>

<span data-ttu-id="01949-111">Um contorno pode ter uma hierarquia, onde grupos menores são aninhados em grupos maiores.</span><span class="sxs-lookup"><span data-stu-id="01949-111">An outline can have a hierarchy, where smaller groups are nested under larger groups.</span></span> <span data-ttu-id="01949-112">Isso permite que o contorno seja exibido em diferentes níveis.</span><span class="sxs-lookup"><span data-stu-id="01949-112">This allows the outline to be viewed at different levels.</span></span> <span data-ttu-id="01949-113">Alterar o nível de contorno visível pode ser feito programaticamente por meio do [método Worksheet.showOutlineLevels.](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)</span><span class="sxs-lookup"><span data-stu-id="01949-113">Changing the visible outline level can be done programmatically through the [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) method.</span></span> <span data-ttu-id="01949-114">Observe que o Excel só dá suporte a oito níveis de grupos de contornos.</span><span class="sxs-lookup"><span data-stu-id="01949-114">Note that Excel only supports eight levels of outline groups.</span></span>

<span data-ttu-id="01949-115">O exemplo de código a seguir cria um contorno com dois níveis de grupos para as linhas e colunas.</span><span class="sxs-lookup"><span data-stu-id="01949-115">The following code sample creates an outline with two levels of groups for both the rows and columns.</span></span> <span data-ttu-id="01949-116">A imagem subsequente mostra os agrupamentos desse contorno.</span><span class="sxs-lookup"><span data-stu-id="01949-116">The subsequent image shows the groupings of that outline.</span></span> <span data-ttu-id="01949-117">No exemplo de código, os intervalos que estão sendo agrupados não incluem a linha ou coluna do controle de contorno (os "Totais" deste exemplo).</span><span class="sxs-lookup"><span data-stu-id="01949-117">In the code sample, the ranges being grouped do not include the row or column of the outline control (the "Totals" for this example).</span></span> <span data-ttu-id="01949-118">Um grupo define o que será recolhido, não a linha ou coluna com o controle.</span><span class="sxs-lookup"><span data-stu-id="01949-118">A group defines what will be collapsed, not the row or column with the control.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);
```

![Um intervalo com um contorno de dois níveis e duas dimensões](../images/excel-outline.png)

## <a name="remove-grouping-from-rows-or-columns-of-a-range"></a><span data-ttu-id="01949-120">Remover o agrupamento de linhas ou colunas de um intervalo</span><span class="sxs-lookup"><span data-stu-id="01949-120">Remove grouping from rows or columns of a range</span></span>

<span data-ttu-id="01949-121">Para desagrupar um grupo de linhas ou colunas, use o [método Range.ungroup.](/javascript/api/excel/excel.range#ungroup-groupoption-)</span><span class="sxs-lookup"><span data-stu-id="01949-121">To ungroup a row or column group, use the [Range.ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) method.</span></span> <span data-ttu-id="01949-122">Isso remove o nível mais externo do contorno.</span><span class="sxs-lookup"><span data-stu-id="01949-122">This removes the outermost level from the outline.</span></span> <span data-ttu-id="01949-123">Se vários grupos do mesmo tipo de linha ou coluna estão no mesmo nível dentro do intervalo especificado, todos esses grupos serão desagrupados.</span><span class="sxs-lookup"><span data-stu-id="01949-123">If multiple groups of the same row or column type are at the same level within the specified range, all of those groups are ungrouped.</span></span>

## <a name="see-also"></a><span data-ttu-id="01949-124">Confira também</span><span class="sxs-lookup"><span data-stu-id="01949-124">See also</span></span>

- [<span data-ttu-id="01949-125">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="01949-125">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="01949-126">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="01949-126">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="01949-127">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="01949-127">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
