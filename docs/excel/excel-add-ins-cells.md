---
title: Trabalhe com células usando a API JavaScript do Excel.
description: Aprenda a definição da API JavaScript do Excel de uma célula e saiba como trabalhar com células.
ms.date: 04/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5fcfeeef52f17c22d13ed3c1a10851f1d8e69204
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652813"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a><span data-ttu-id="d2561-103">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="d2561-103">Work with cells using the Excel JavaScript API</span></span>

<span data-ttu-id="d2561-104">A API JavaScript do Excel não tem um objeto ou classe "Cell".</span><span class="sxs-lookup"><span data-stu-id="d2561-104">The Excel JavaScript API doesn't have a "Cell" object or class.</span></span> <span data-ttu-id="d2561-105">Em vez disso, todas as células do Excel são `Range` objetos.</span><span class="sxs-lookup"><span data-stu-id="d2561-105">Instead, all Excel cells are `Range` objects.</span></span> <span data-ttu-id="d2561-106">Uma célula individual na interface do usuário do Excel é traduzida para um `Range` objeto com uma célula na API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="d2561-106">An individual cell in the Excel UI translates to a `Range` object with one cell in the Excel JavaScript API.</span></span>

<span data-ttu-id="d2561-107">Um `Range` objeto também pode conter várias células contíguas.</span><span class="sxs-lookup"><span data-stu-id="d2561-107">A `Range` object can also contain multiple, contiguous cells.</span></span> <span data-ttu-id="d2561-108">Células contíguas formam um retângulo ininterrupto (incluindo linhas ou colunas simples).</span><span class="sxs-lookup"><span data-stu-id="d2561-108">Contiguous cells form an unbroken rectangle (including single rows or columns).</span></span> <span data-ttu-id="d2561-109">Para saber mais sobre como trabalhar com células que não são contíguas, consulte Trabalhar com células [descontíguas usando o objeto RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object).</span><span class="sxs-lookup"><span data-stu-id="d2561-109">To learn about working with cells that are not contiguous, see [Work with discontiguous cells using the RangeAreas object](#work-with-discontiguous-cells-using-the-rangeareas-object).</span></span>

<span data-ttu-id="d2561-110">Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="d2561-110">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="excel-javascript-apis-that-mention-cells"></a><span data-ttu-id="d2561-111">APIs JavaScript do Excel que mencionam células</span><span class="sxs-lookup"><span data-stu-id="d2561-111">Excel JavaScript APIs that mention cells</span></span>

<span data-ttu-id="d2561-112">Mesmo que a API JavaScript do Excel não tenha um objeto ou classe "Cell", vários nomes de API mencionam células.</span><span class="sxs-lookup"><span data-stu-id="d2561-112">Even though the Excel JavaScript API doesn't have a "Cell" object or class, a number of API names mention cells.</span></span> <span data-ttu-id="d2561-113">Essas APIs controlam propriedades de célula, como cor, formatação de texto e fonte.</span><span class="sxs-lookup"><span data-stu-id="d2561-113">These APIs control cell properties like color, text formatting, and font.</span></span>

<span data-ttu-id="d2561-114">A lista a seguir das APIs JavaScript do Excel referem-se a células.</span><span class="sxs-lookup"><span data-stu-id="d2561-114">The following list of Excel JavaScript APIs refer to cells.</span></span>

- [<span data-ttu-id="d2561-115">CellBorder</span><span class="sxs-lookup"><span data-stu-id="d2561-115">CellBorder</span></span>](/javascript/api/excel/excel.cellborder)
- [<span data-ttu-id="d2561-116">CellBorderCollection</span><span class="sxs-lookup"><span data-stu-id="d2561-116">CellBorderCollection</span></span>](/javascript/api/excel/excel.cellbordercollection)
- [<span data-ttu-id="d2561-117">CellProperties</span><span class="sxs-lookup"><span data-stu-id="d2561-117">CellProperties</span></span>](/javascript/api/excel/excel.cellproperties)
- [<span data-ttu-id="d2561-118">CellPropertiesFill</span><span class="sxs-lookup"><span data-stu-id="d2561-118">CellPropertiesFill</span></span>](/javascript/api/excel/excel.cellpropertiesfill)
- [<span data-ttu-id="d2561-119">CellPropertiesFont</span><span class="sxs-lookup"><span data-stu-id="d2561-119">CellPropertiesFont</span></span>](/javascript/api/excel/excel.cellpropertiesfont)
- [<span data-ttu-id="d2561-120">CellPropertiesFormat</span><span class="sxs-lookup"><span data-stu-id="d2561-120">CellPropertiesFormat</span></span>](/javascript/api/excel/excel.cellpropertiesformat)
- [<span data-ttu-id="d2561-121">CellPropertiesProtection</span><span class="sxs-lookup"><span data-stu-id="d2561-121">CellPropertiesProtection</span></span>](/javascript/api/excel/excel.cellpropertiesprotection)
- [<span data-ttu-id="d2561-122">CellValueConditionalFormat</span><span class="sxs-lookup"><span data-stu-id="d2561-122">CellValueConditionalFormat</span></span>](/javascript/api/excel/excel.cellvalueconditionalformat)
- [<span data-ttu-id="d2561-123">ConditionalCellValueRule</span><span class="sxs-lookup"><span data-stu-id="d2561-123">ConditionalCellValueRule</span></span>](/javascript/api/excel/excel.conditionalcellvaluerule)
- [<span data-ttu-id="d2561-124">SettableCellProperties</span><span class="sxs-lookup"><span data-stu-id="d2561-124">SettableCellProperties</span></span>](/javascript/api/excel/excel.settablecellproperties)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a><span data-ttu-id="d2561-125">Trabalhar com células desconsiguadas usando o objeto RangeAreas</span><span class="sxs-lookup"><span data-stu-id="d2561-125">Work with discontiguous cells using the RangeAreas object</span></span>

<span data-ttu-id="d2561-126">O [objeto RangeAreas](/javascript/api/excel/excel.rangeareas) permite que o seu complemento execute operações em vários intervalos de uma só vez.</span><span class="sxs-lookup"><span data-stu-id="d2561-126">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="d2561-127">Esses intervalos podem ser contíguos, mas não precisam ser.</span><span class="sxs-lookup"><span data-stu-id="d2561-127">These ranges may be contiguous, but they don't have to be.</span></span> <span data-ttu-id="d2561-128">`RangeAreas` são descritas ainda mais no artigo [Trabalhar com vários intervalos simultaneamente em suplementos do Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="d2561-128">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d2561-129">Confira também</span><span class="sxs-lookup"><span data-stu-id="d2561-129">See also</span></span>

- [<span data-ttu-id="d2561-130">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="d2561-130">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="d2561-131">Obter um intervalo usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="d2561-131">Get a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-get.md)
- [<span data-ttu-id="d2561-132">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="d2561-132">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
