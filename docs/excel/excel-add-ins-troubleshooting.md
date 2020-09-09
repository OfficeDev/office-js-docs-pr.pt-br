---
title: Solucionando problemas de suplementos do Excel
description: Saiba como solucionar erros de desenvolvimento em suplementos do Excel.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 1bdd96772d3a221ca3a02e3d5dfcfa16561dd5f1
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409372"
---
# <a name="troubleshooting-excel-add-ins"></a><span data-ttu-id="679cf-103">Solucionando problemas de suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="679cf-103">Troubleshooting Excel Add-ins</span></span>

<span data-ttu-id="679cf-104">Este artigo discute a solução de problemas exclusivos para o Excel.</span><span class="sxs-lookup"><span data-stu-id="679cf-104">This article discusses troubleshooting issues that are unique to Excel.</span></span> <span data-ttu-id="679cf-105">Use a ferramenta de comentários na parte inferior da página para sugerir outros problemas que podem ser adicionados ao artigo.</span><span class="sxs-lookup"><span data-stu-id="679cf-105">Please use the feedback tool at the bottom of the page to suggest other issues that can be added to the article.</span></span>

## <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="679cf-106">Limitações de API quando a pasta de trabalho ativa alterna</span><span class="sxs-lookup"><span data-stu-id="679cf-106">API limitations when the active workbook switches</span></span>

<span data-ttu-id="679cf-107">Os suplementos para Excel se destinam a operar em uma única pasta de trabalho por vez.</span><span class="sxs-lookup"><span data-stu-id="679cf-107">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="679cf-108">Os erros podem ocorrer quando uma pasta de trabalho separada da que está executando o suplemento Obtém o foco.</span><span class="sxs-lookup"><span data-stu-id="679cf-108">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="679cf-109">Isso ocorre apenas quando determinados métodos estão no processo de chamada quando o foco é alterado.</span><span class="sxs-lookup"><span data-stu-id="679cf-109">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="679cf-110">As seguintes APIs são afetadas por essa opção de pasta de trabalho:</span><span class="sxs-lookup"><span data-stu-id="679cf-110">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="679cf-111">API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="679cf-111">Excel JavaScript API</span></span> | <span data-ttu-id="679cf-112">Erro gerado</span><span class="sxs-lookup"><span data-stu-id="679cf-112">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="679cf-113">GeneralException</span><span class="sxs-lookup"><span data-stu-id="679cf-113">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="679cf-114">GeneralException</span><span class="sxs-lookup"><span data-stu-id="679cf-114">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="679cf-115">GeneralException</span><span class="sxs-lookup"><span data-stu-id="679cf-115">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="679cf-116">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="679cf-116">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="679cf-117">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="679cf-117">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="679cf-118">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="679cf-118">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="679cf-119">GeneralException</span><span class="sxs-lookup"><span data-stu-id="679cf-119">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="679cf-120">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="679cf-120">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="679cf-121">GeneralException</span><span class="sxs-lookup"><span data-stu-id="679cf-121">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="679cf-122">GeneralException</span><span class="sxs-lookup"><span data-stu-id="679cf-122">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="679cf-123">GeneralException</span><span class="sxs-lookup"><span data-stu-id="679cf-123">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="679cf-124">GeneralException</span><span class="sxs-lookup"><span data-stu-id="679cf-124">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="679cf-125">GeneralException</span><span class="sxs-lookup"><span data-stu-id="679cf-125">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="679cf-126">GeneralException</span><span class="sxs-lookup"><span data-stu-id="679cf-126">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="679cf-127">GeneralException</span><span class="sxs-lookup"><span data-stu-id="679cf-127">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="679cf-128">GeneralException</span><span class="sxs-lookup"><span data-stu-id="679cf-128">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="679cf-129">Isso aplica-se apenas a várias pastas de trabalho do Excel abertas no Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="679cf-129">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="679cf-130">Coautoria</span><span class="sxs-lookup"><span data-stu-id="679cf-130">Coauthoring</span></span>

<span data-ttu-id="679cf-131">Veja [coautoria em suplementos do Excel](co-authoring-in-excel-add-ins.md) para padrões a serem usados com eventos em um ambiente de coautoria.</span><span class="sxs-lookup"><span data-stu-id="679cf-131">See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="679cf-132">O artigo também aborda possíveis conflitos de mesclagem ao usar determinadas APIs, como [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .</span><span class="sxs-lookup"><span data-stu-id="679cf-132">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="see-also"></a><span data-ttu-id="679cf-133">Confira também</span><span class="sxs-lookup"><span data-stu-id="679cf-133">See also</span></span>

- [<span data-ttu-id="679cf-134">Solucionar erros de desenvolvimento com suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="679cf-134">Troubleshoot development errors with Office Add-ins</span></span>](../testing/troubleshoot-development-errors.md)
- [<span data-ttu-id="679cf-135">Solucionar erros de usuários com Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="679cf-135">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
