---
title: Visão geral da API JavaScript do Excel
description: ''
ms.date: 07/05/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: e6064bf7e7dce6931079fc2d3eb262533da7edf3
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575629"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="34c21-102">Visão geral da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="34c21-102">Excel JavaScript API overview</span></span>

<span data-ttu-id="34c21-103">Um suplemento do Excel interage com objetos no Excel usando a API JavaScript para Office, que inclui dois modelos de objeto JavaScript:</span><span class="sxs-lookup"><span data-stu-id="34c21-103">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="34c21-104">**API JavaScript do Excel**: introduzida com o Office 2016, a [API JavaScript do Excel](/javascript/api/excel) fornece objetos fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais.</span><span class="sxs-lookup"><span data-stu-id="34c21-104">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="34c21-105">**APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="34c21-105">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="34c21-106">Esta seção da documentação concentra-se na API JavaScript do Excel, que você usará para desenvolver a maior parte da funcionalidade em suplementos direcionados para o Excel na Web ou para o Excel 2016 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="34c21-106">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="34c21-107">Para saber mais sobre a API comum, confira [API do JavaScript para Office](../javascript-api-for-office.md).</span><span class="sxs-lookup"><span data-stu-id="34c21-107">For more information about the distinction between host-specific APIs and Common APIs, see [JavaScript API for Office](../javascript-api-for-office.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="34c21-108">Aprenda conceitos de programação</span><span class="sxs-lookup"><span data-stu-id="34c21-108">Learn programming concepts</span></span>

<span data-ttu-id="34c21-109">Confira os artigos a seguir para obter informações sobre conceitos de programação importantes:</span><span class="sxs-lookup"><span data-stu-id="34c21-109">See the following articles for information about important programming concepts:</span></span>
 
- [<span data-ttu-id="34c21-110">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="34c21-110">Fundamental programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-core-concepts.md)

- [<span data-ttu-id="34c21-111">Conceitos avançados de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="34c21-111">Advanced programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-advanced-concepts.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="34c21-112">Saiba mais sobre recursos da API</span><span class="sxs-lookup"><span data-stu-id="34c21-112">Learn about API capabilities</span></span>

<span data-ttu-id="34c21-113">Use outros artigos nesta seção da documentação para saber mais sobre como trabalhar com [eventos](../../excel/excel-add-ins-events.md), [gráficos](../../excel/excel-add-ins-charts.md), [intervalos](../../excel/excel-add-ins-ranges.md), [tabelas](../../excel/excel-add-ins-tables.md) [planilhas](../../excel/excel-add-ins-worksheets.md), e muito mais.</span><span class="sxs-lookup"><span data-stu-id="34c21-113">Use other articles in this section of the documentation to learn about working with [events](../../excel/excel-add-ins-events.md), [charts](../../excel/excel-add-ins-charts.md), [ranges](../../excel/excel-add-ins-ranges.md), [tables](../../excel/excel-add-ins-tables.md), [worksheets](../../excel/excel-add-ins-worksheets.md), and more.</span></span> <span data-ttu-id="34c21-114">Além disso, nesta seção, você encontrará instruções sobre os conceitos da API JavaScript do Excel, como [coautoria em suplementos do Excel](../../excel/co-authoring-in-excel-add-ins.md), [validação de dados](../../excel/excel-add-ins-data-validation.md), [tratamento de erros](../../excel/excel-add-ins-error-handling.md) e [otimização de desempenho](../../excel/performance.md).</span><span class="sxs-lookup"><span data-stu-id="34c21-114">Also in this section, you'll find guidance about Excel JavaScript API concepts such as [coauthoring in Excel add-ins](../../excel/co-authoring-in-excel-add-ins.md), [data validation](../../excel/excel-add-ins-data-validation.md), [error handling](../../excel/excel-add-ins-error-handling.md), and [performance optimization](../../excel/performance.md).</span></span> <span data-ttu-id="34c21-115">Confira o Sumário para obter a lista completa de artigos disponíveis.</span><span class="sxs-lookup"><span data-stu-id="34c21-115">See the table of contents for the complete list of available articles.</span></span>

<span data-ttu-id="34c21-116">Para ter a experiência prática com o uso da API JavaScript do Excel para acessar objetos no Excel, conclua o [tutorial do suplemento do Excel](../../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="34c21-116">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span> 

<span data-ttu-id="34c21-117">Para saber mais sobre o modelo de objeto API JavaScript do Excel, consulte a [Documentação de referência da API JavaScript do Excel](/javascript/api/excel).</span><span class="sxs-lookup"><span data-stu-id="34c21-117">For detailed information about the Excel JavaScript API, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="34c21-118">Experimente amostras de código no Script Lab</span><span class="sxs-lookup"><span data-stu-id="34c21-118">Try out code samples in Script Lab</span></span>

<span data-ttu-id="34c21-119">Use o [Script Lab](../../overview/explore-with-script-lab.md) para começar a trabalhar rapidamente com um conjunto de exemplos internos que mostram como concluir tarefas com a API.</span><span class="sxs-lookup"><span data-stu-id="34c21-119">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="34c21-120">Você pode executar as amostras no Script Lab para ver instantaneamente o resultado no painel de tarefas ou planilha, examinar os exemplos para saber como a API funciona e até mesmo usar amostras para criar um protótipo do seu próprio suplemento.</span><span class="sxs-lookup"><span data-stu-id="34c21-120">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="34c21-121">Confira também</span><span class="sxs-lookup"><span data-stu-id="34c21-121">See also</span></span>

- [<span data-ttu-id="34c21-122">Documentação de Suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="34c21-122">Excel add-ins documentation</span></span>](../../excel/index.md)
- [<span data-ttu-id="34c21-123">Visão geral dos suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="34c21-123">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
- [<span data-ttu-id="34c21-124">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="34c21-124">Excel JavaScript API reference</span></span>](/javascript/api/excel)
- [<span data-ttu-id="34c21-125">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="34c21-125">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
- [<span data-ttu-id="34c21-126">Especificações abertas da API</span><span class="sxs-lookup"><span data-stu-id="34c21-126">API open specifications</span></span>](../openspec/openspec.md)
