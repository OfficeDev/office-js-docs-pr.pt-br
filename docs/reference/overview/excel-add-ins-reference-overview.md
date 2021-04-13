---
title: Visão geral da API JavaScript do Excel
description: Saiba mais sobre as APIs JavaScript do Excel
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 80340b4990b56b2ba4d51f2a028480af3e267828
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51650804"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="2d7bb-103">Visão geral da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="2d7bb-103">Excel JavaScript API overview</span></span>

<span data-ttu-id="2d7bb-104">Um suplemento do Excel interage com objetos no Excel usando a API JavaScript do Office, que inclui dois modelos de objetos JavaScript:</span><span class="sxs-lookup"><span data-stu-id="2d7bb-104">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="2d7bb-105">**API de JavaScript do Excel para**: estas são as [APIs específicas do aplicativo](../../develop/application-specific-api-model.md) para o Excel.</span><span class="sxs-lookup"><span data-stu-id="2d7bb-105">**Excel JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for Excel.</span></span> <span data-ttu-id="2d7bb-106">Introduzida com o Office 2016, a [API de JavaScript do Excel](/javascript/api/excel) fornece objetos fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais.</span><span class="sxs-lookup"><span data-stu-id="2d7bb-106">Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="2d7bb-107">**APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="2d7bb-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="2d7bb-108">Esta seção da documentação concentra-se na API JavaScript do Excel, que você usará para desenvolver a maior parte da funcionalidade em suplementos direcionados para o Excel na Web ou para o Excel 2016 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="2d7bb-108">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="2d7bb-109">Para obter mais informações do API comum, consulte [Modelo do objeto do JavaScript API comum](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="2d7bb-109">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span>

## <a name="learn-object-model-concepts"></a><span data-ttu-id="2d7bb-110">Aprender os conceitos do modelo de objeto</span><span class="sxs-lookup"><span data-stu-id="2d7bb-110">Learn object model concepts</span></span>

<span data-ttu-id="2d7bb-111">Confira o [Modelo de objeto JavaScript do Excel em suplementos do Office](../../excel/excel-add-ins-core-concepts.md) para obter informações sobre conceitos importantes do modelo de objeto.</span><span class="sxs-lookup"><span data-stu-id="2d7bb-111">See [Excel JavaScript object model in Office Add-ins](../../excel/excel-add-ins-core-concepts.md) for information about important object model concepts.</span></span>

<span data-ttu-id="2d7bb-112">Para ter a experiência prática com o uso da API de JavaScript do Excel para acessar objetos no Excel, conclua o [Tutorial do suplemento do Excel](../../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="2d7bb-112">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span>

## <a name="learn-api-capabilities"></a><span data-ttu-id="2d7bb-113">Conheça os recursos da API</span><span class="sxs-lookup"><span data-stu-id="2d7bb-113">Learn API capabilities</span></span>

<span data-ttu-id="2d7bb-114">Cada recurso principal da API do Excel possui um artigo ou conjunto de artigos explorando o que esse recurso pode fazer e o modelo de objeto relevante.</span><span class="sxs-lookup"><span data-stu-id="2d7bb-114">Each major Excel API feature has an article or set of articles exploring what that feature can do and the relevant object model.</span></span>

* [<span data-ttu-id="2d7bb-115">Gráficos</span><span class="sxs-lookup"><span data-stu-id="2d7bb-115">Charts</span></span>](../../excel/excel-add-ins-charts.md)
* [<span data-ttu-id="2d7bb-116">Comentário</span><span class="sxs-lookup"><span data-stu-id="2d7bb-116">Comments</span></span>](../../excel/excel-add-ins-comments.md)
* [<span data-ttu-id="2d7bb-117">Formatação condicional</span><span class="sxs-lookup"><span data-stu-id="2d7bb-117">Conditional formatting</span></span>](../../excel/excel-add-ins-conditional-formatting.md)
* [<span data-ttu-id="2d7bb-118">Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="2d7bb-118">Custom functions</span></span>](../../excel/custom-functions-overview.md)
* [<span data-ttu-id="2d7bb-119">Validação de dados</span><span class="sxs-lookup"><span data-stu-id="2d7bb-119">Data validation</span></span>](../../excel/excel-add-ins-data-validation.md)
* [<span data-ttu-id="2d7bb-120">Eventos</span><span class="sxs-lookup"><span data-stu-id="2d7bb-120">Events</span></span>](../../excel/excel-add-ins-events.md)
* [<span data-ttu-id="2d7bb-121">Tabelas Dinâmicas</span><span class="sxs-lookup"><span data-stu-id="2d7bb-121">PivotTables</span></span>](../../excel/excel-add-ins-pivottables.md)
* <span data-ttu-id="2d7bb-122">[Faixas](../../excel/excel-add-ins-ranges-get.md) e [Células](../../excel/excel-add-ins-cells.md)</span><span class="sxs-lookup"><span data-stu-id="2d7bb-122">[Ranges](../../excel/excel-add-ins-ranges-get.md) and [Cells](../../excel/excel-add-ins-cells.md)</span></span>
* [<span data-ttu-id="2d7bb-123">RangeAreas (vários intervalos)</span><span class="sxs-lookup"><span data-stu-id="2d7bb-123">RangeAreas (Multiple ranges)</span></span>](../../excel/excel-add-ins-multiple-ranges.md)
* [<span data-ttu-id="2d7bb-124">Formas</span><span class="sxs-lookup"><span data-stu-id="2d7bb-124">Shapes</span></span>](../../excel/excel-add-ins-shapes.md)
* [<span data-ttu-id="2d7bb-125">Tabelas</span><span class="sxs-lookup"><span data-stu-id="2d7bb-125">Tables</span></span>](../../excel/excel-add-ins-tables.md)
* [<span data-ttu-id="2d7bb-126">Pastas de trabalho e APIs no Nível do Aplicativo</span><span class="sxs-lookup"><span data-stu-id="2d7bb-126">Workbooks and Application-level APIs</span></span>](../../excel/excel-add-ins-workbooks.md)
* [<span data-ttu-id="2d7bb-127">Planilhas</span><span class="sxs-lookup"><span data-stu-id="2d7bb-127">Worksheets</span></span>](../../excel/excel-add-ins-worksheets.md)

<span data-ttu-id="2d7bb-128">Para saber mais sobre o modelo de objeto API JavaScript do Excel, consulte a [Documentação de referência da API JavaScript do Excel](/javascript/api/excel).</span><span class="sxs-lookup"><span data-stu-id="2d7bb-128">For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="2d7bb-129">Experimente amostras de código no Script Lab</span><span class="sxs-lookup"><span data-stu-id="2d7bb-129">Try out code samples in Script Lab</span></span>

<span data-ttu-id="2d7bb-130">Use o [Script Lab](../../overview/explore-with-script-lab.md) para começar a trabalhar rapidamente com um conjunto de exemplos internos que mostram como concluir tarefas com a API.</span><span class="sxs-lookup"><span data-stu-id="2d7bb-130">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="2d7bb-131">Você pode executar as amostras no Script Lab para ver instantaneamente o resultado no painel de tarefas ou planilha, examinar os exemplos para saber como a API funciona e até mesmo usar amostras para criar um protótipo do seu próprio suplemento.</span><span class="sxs-lookup"><span data-stu-id="2d7bb-131">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="2d7bb-132">Confira também</span><span class="sxs-lookup"><span data-stu-id="2d7bb-132">See also</span></span>

* [<span data-ttu-id="2d7bb-133">Documentação de Suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="2d7bb-133">Excel add-ins documentation</span></span>](../../excel/index.yml)
* [<span data-ttu-id="2d7bb-134">Visão geral dos suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="2d7bb-134">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
* [<span data-ttu-id="2d7bb-135">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="2d7bb-135">Excel JavaScript API reference</span></span>](/javascript/api/excel)
* [<span data-ttu-id="2d7bb-136">Disponibilidade de aplicativos e plataformas de cliente Office para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="2d7bb-136">Office client application and platform availability for Office Add-ins</span></span>](../../overview/office-add-in-availability.md)
* [<span data-ttu-id="2d7bb-137">Usando o modelo de API específica do aplicativo</span><span class="sxs-lookup"><span data-stu-id="2d7bb-137">Using the application-specific API model</span></span>](../../develop/application-specific-api-model.md)
