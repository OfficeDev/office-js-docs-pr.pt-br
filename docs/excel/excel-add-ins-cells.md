---
title: Trabalhe com células usando a API JavaScript do Excel.
description: Aprenda a definição da API JavaScript do Excel de uma célula e saiba como trabalhar com células.
ms.date: 04/16/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad8ca985b6bbdcf19920c36c371e690f61639f16
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917097"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a><span data-ttu-id="545ea-103">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="545ea-103">Work with cells using the Excel JavaScript API</span></span>

<span data-ttu-id="545ea-104">A API JavaScript do Excel não tem um objeto ou classe "Célula".</span><span class="sxs-lookup"><span data-stu-id="545ea-104">The Excel JavaScript API doesn't have a "Cell" object or class.</span></span> <span data-ttu-id="545ea-105">Em vez disso, todas as células do Excel são `Range` objetos.</span><span class="sxs-lookup"><span data-stu-id="545ea-105">Instead, all Excel cells are `Range` objects.</span></span> <span data-ttu-id="545ea-106">Uma célula individual na interface do usuário do Excel se traduz em um objeto `Range` com uma célula na API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="545ea-106">An individual cell in the Excel UI translates to a `Range` object with one cell in the Excel JavaScript API.</span></span>

<span data-ttu-id="545ea-107">Um `Range` objeto também pode conter várias células contíguas.</span><span class="sxs-lookup"><span data-stu-id="545ea-107">A `Range` object can also contain multiple, contiguous cells.</span></span> <span data-ttu-id="545ea-108">Células contíguas formam um retângulo ininterrupto (incluindo linhas ou colunas simples).</span><span class="sxs-lookup"><span data-stu-id="545ea-108">Contiguous cells form an unbroken rectangle (including single rows or columns).</span></span> <span data-ttu-id="545ea-109">Para saber mais sobre como trabalhar com células que não são contíguas, consulte Trabalhar com células [descontíguas usando o objeto RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object).</span><span class="sxs-lookup"><span data-stu-id="545ea-109">To learn about working with cells that are not contiguous, see [Work with discontiguous cells using the RangeAreas object](#work-with-discontiguous-cells-using-the-rangeareas-object).</span></span>

<span data-ttu-id="545ea-110">Para ver a lista completa de propriedades e métodos compatíveis com o objeto, consulte `Range` [Range Object (API JavaScript para Excel)](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="545ea-110">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a><span data-ttu-id="545ea-111">Trabalhar com células desconsiguadas usando o objeto RangeAreas</span><span class="sxs-lookup"><span data-stu-id="545ea-111">Work with discontiguous cells using the RangeAreas object</span></span>

<span data-ttu-id="545ea-112">O [objeto RangeAreas](/javascript/api/excel/excel.rangeareas) permite que o seu complemento execute operações em vários intervalos de uma só vez.</span><span class="sxs-lookup"><span data-stu-id="545ea-112">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="545ea-113">Esses intervalos podem ser contíguos, mas não precisam ser.</span><span class="sxs-lookup"><span data-stu-id="545ea-113">These ranges may be contiguous, but they don't have to be.</span></span> <span data-ttu-id="545ea-114">`RangeAreas` são descritas ainda mais no artigo [Trabalhar com vários intervalos simultaneamente em suplementos do Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="545ea-114">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="545ea-115">Confira também</span><span class="sxs-lookup"><span data-stu-id="545ea-115">See also</span></span>

- [<span data-ttu-id="545ea-116">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="545ea-116">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="545ea-117">Obter um intervalo usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="545ea-117">Get a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-get.md)
- [<span data-ttu-id="545ea-118">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="545ea-118">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
