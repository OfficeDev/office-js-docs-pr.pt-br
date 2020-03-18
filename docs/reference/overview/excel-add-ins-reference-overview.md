---
title: Visão geral da API JavaScript do Excel
description: Saiba mais sobre as APIs JavaScript do Excel
ms.date: 02/19/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: d4721c349de7b93ec34ce3c35fbcee25438564db
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719914"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="12385-103">Visão geral da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="12385-103">Excel JavaScript API overview</span></span>

<span data-ttu-id="12385-104">Um suplemento do Excel interage com objetos no Excel usando a API JavaScript do Office, que inclui dois modelos de objetos JavaScript:</span><span class="sxs-lookup"><span data-stu-id="12385-104">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="12385-105">**API JavaScript do Excel**: introduzida com o Office 2016, a [API JavaScript do Excel](/javascript/api/excel) fornece objetos fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais.</span><span class="sxs-lookup"><span data-stu-id="12385-105">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="12385-106">**APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="12385-106">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="12385-107">Esta seção da documentação concentra-se na API JavaScript do Excel, que você usará para desenvolver a maior parte da funcionalidade em suplementos direcionados para o Excel na Web ou para o Excel 2016 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="12385-107">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="12385-108">Para obter mais informações do API comum, consulte [Modelo do objeto do JavaScript API comum](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="12385-108">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="12385-109">Aprenda conceitos de programação</span><span class="sxs-lookup"><span data-stu-id="12385-109">Learn programming concepts</span></span>

<span data-ttu-id="12385-110">Confira os artigos a seguir para obter informações sobre conceitos de programação importantes:</span><span class="sxs-lookup"><span data-stu-id="12385-110">See the following articles for information about important programming concepts:</span></span>
 
- [<span data-ttu-id="12385-111">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="12385-111">Fundamental programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-core-concepts.md)

- [<span data-ttu-id="12385-112">Conceitos avançados de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="12385-112">Advanced programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-advanced-concepts.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="12385-113">Saiba mais sobre recursos da API</span><span class="sxs-lookup"><span data-stu-id="12385-113">Learn about API capabilities</span></span>

<span data-ttu-id="12385-114">Use outros artigos nesta seção da documentação para saber mais sobre como trabalhar com [eventos](../../excel/excel-add-ins-events.md), [gráficos](../../excel/excel-add-ins-charts.md), [intervalos](../../excel/excel-add-ins-ranges.md), [tabelas](../../excel/excel-add-ins-tables.md) [planilhas](../../excel/excel-add-ins-worksheets.md), e muito mais.</span><span class="sxs-lookup"><span data-stu-id="12385-114">Use other articles in this section of the documentation to learn about working with [events](../../excel/excel-add-ins-events.md), [charts](../../excel/excel-add-ins-charts.md), [ranges](../../excel/excel-add-ins-ranges.md), [tables](../../excel/excel-add-ins-tables.md), [worksheets](../../excel/excel-add-ins-worksheets.md), and more.</span></span> <span data-ttu-id="12385-115">Além disso, nesta seção, você encontrará instruções sobre os conceitos da API JavaScript do Excel, como [coautoria em suplementos do Excel](../../excel/co-authoring-in-excel-add-ins.md), [validação de dados](../../excel/excel-add-ins-data-validation.md), [tratamento de erros](../../excel/excel-add-ins-error-handling.md) e [otimização de desempenho](../../excel/performance.md).</span><span class="sxs-lookup"><span data-stu-id="12385-115">Also in this section, you'll find guidance about Excel JavaScript API concepts such as [coauthoring in Excel add-ins](../../excel/co-authoring-in-excel-add-ins.md), [data validation](../../excel/excel-add-ins-data-validation.md), [error handling](../../excel/excel-add-ins-error-handling.md), and [performance optimization](../../excel/performance.md).</span></span> <span data-ttu-id="12385-116">Confira o Sumário para obter a lista completa de artigos disponíveis.</span><span class="sxs-lookup"><span data-stu-id="12385-116">See the table of contents for the complete list of available articles.</span></span>

<span data-ttu-id="12385-117">Para ter a experiência prática com o uso da API JavaScript do Excel para acessar objetos no Excel, conclua o [tutorial do suplemento do Excel](../../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="12385-117">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span> 

<span data-ttu-id="12385-118">Para saber mais sobre o modelo de objeto API JavaScript do Excel, consulte a [Documentação de referência da API JavaScript do Excel](/javascript/api/excel).</span><span class="sxs-lookup"><span data-stu-id="12385-118">For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="12385-119">Experimente amostras de código no Script Lab</span><span class="sxs-lookup"><span data-stu-id="12385-119">Try out code samples in Script Lab</span></span>

<span data-ttu-id="12385-120">Use o [Script Lab](../../overview/explore-with-script-lab.md) para começar a trabalhar rapidamente com um conjunto de exemplos internos que mostram como concluir tarefas com a API.</span><span class="sxs-lookup"><span data-stu-id="12385-120">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="12385-121">Você pode executar as amostras no Script Lab para ver instantaneamente o resultado no painel de tarefas ou planilha, examinar os exemplos para saber como a API funciona e até mesmo usar amostras para criar um protótipo do seu próprio suplemento.</span><span class="sxs-lookup"><span data-stu-id="12385-121">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="12385-122">Confira também</span><span class="sxs-lookup"><span data-stu-id="12385-122">See also</span></span>

- [<span data-ttu-id="12385-123">Documentação de Suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="12385-123">Excel add-ins documentation</span></span>](../../excel/index.md)
- [<span data-ttu-id="12385-124">Visão geral dos suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="12385-124">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
- [<span data-ttu-id="12385-125">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="12385-125">Excel JavaScript API reference</span></span>](/javascript/api/excel)
- [<span data-ttu-id="12385-126">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="12385-126">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
- [<span data-ttu-id="12385-127">Especificações abertas da API</span><span class="sxs-lookup"><span data-stu-id="12385-127">API open specifications</span></span>](../openspec/openspec.md)
