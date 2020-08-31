---
title: Visão geral da API JavaScript do Excel
description: Saiba mais sobre as APIs JavaScript do Excel
ms.date: 07/28/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: e589bd7ce814211759cc731d828e9c180339ea1f
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293657"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="ee6e6-103">Visão geral da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="ee6e6-103">Excel JavaScript API overview</span></span>

<span data-ttu-id="ee6e6-104">Um suplemento do Excel interage com objetos no Excel usando a API JavaScript do Office, que inclui dois modelos de objetos JavaScript:</span><span class="sxs-lookup"><span data-stu-id="ee6e6-104">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="ee6e6-105">**API de JavaScript do Excel para**: estas são as [APIs específicas do aplicativo](../../develop/application-specific-api-model.md) para o Excel.</span><span class="sxs-lookup"><span data-stu-id="ee6e6-105">**Excel JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for Excel.</span></span> <span data-ttu-id="ee6e6-106">Introduzida com o Office 2016, a [API de JavaScript do Excel](/javascript/api/excel) fornece objetos fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais.</span><span class="sxs-lookup"><span data-stu-id="ee6e6-106">Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="ee6e6-107">**APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="ee6e6-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="ee6e6-108">Esta seção da documentação concentra-se na API JavaScript do Excel, que você usará para desenvolver a maior parte da funcionalidade em suplementos direcionados para o Excel na Web ou para o Excel 2016 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="ee6e6-108">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="ee6e6-109">Para obter mais informações do API comum, consulte [Modelo do objeto do JavaScript API comum](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="ee6e6-109">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span>

## <a name="learn-programming-concepts"></a><span data-ttu-id="ee6e6-110">Aprenda conceitos de programação</span><span class="sxs-lookup"><span data-stu-id="ee6e6-110">Learn programming concepts</span></span>

<span data-ttu-id="ee6e6-111">Veja [Conceitos fundamentais de programação com a API de JavaScript do Excel](../../excel/excel-add-ins-core-concepts.md) para obter informações sobre conceitos de programação importantes.</span><span class="sxs-lookup"><span data-stu-id="ee6e6-111">See [Fundamental programming concepts with the Excel JavaScript API](../../excel/excel-add-ins-core-concepts.md) for information about important programming concepts.</span></span>

<span data-ttu-id="ee6e6-112">Para ter a experiência prática com o uso da API de JavaScript do Excel para acessar objetos no Excel, conclua o [Tutorial do suplemento do Excel](../../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="ee6e6-112">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span>

## <a name="learn-api-capabilities"></a><span data-ttu-id="ee6e6-113">Conheça os recursos da API</span><span class="sxs-lookup"><span data-stu-id="ee6e6-113">Learn API capabilities</span></span>

<span data-ttu-id="ee6e6-114">Cada recurso principal da API do Excel tem um artigo explorando o que pode ser feito e o modelo de objeto relevante.</span><span class="sxs-lookup"><span data-stu-id="ee6e6-114">Each major Excel API feature has an article exploring what that feature can do and the relevant object model.</span></span>

* [<span data-ttu-id="ee6e6-115">Gráficos</span><span class="sxs-lookup"><span data-stu-id="ee6e6-115">Charts</span></span>](../../excel/excel-add-ins-charts.md)
* [<span data-ttu-id="ee6e6-116">Comentário</span><span class="sxs-lookup"><span data-stu-id="ee6e6-116">Comments</span></span>](../../excel/excel-add-ins-comments.md)
* [<span data-ttu-id="ee6e6-117">Formatação condicional</span><span class="sxs-lookup"><span data-stu-id="ee6e6-117">Conditional formatting</span></span>](../../excel/excel-add-ins-conditional-formatting.md)
* [<span data-ttu-id="ee6e6-118">Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ee6e6-118">Custom functions</span></span>](../../excel/custom-functions-overview.md)
* [<span data-ttu-id="ee6e6-119">Validação de dados</span><span class="sxs-lookup"><span data-stu-id="ee6e6-119">Data validation</span></span>](../../excel/excel-add-ins-data-validation.md)
* [<span data-ttu-id="ee6e6-120">Eventos</span><span class="sxs-lookup"><span data-stu-id="ee6e6-120">Events</span></span>](../../excel/excel-add-ins-events.md)
* [<span data-ttu-id="ee6e6-121">Vários intervalos (RangeArea)</span><span class="sxs-lookup"><span data-stu-id="ee6e6-121">Multiple ranges (RangeArea)</span></span>](../../excel/excel-add-ins-multiple-ranges.md)
* [<span data-ttu-id="ee6e6-122">PivotTables</span><span class="sxs-lookup"><span data-stu-id="ee6e6-122">PivotTables</span></span>](../../excel/excel-add-ins-pivottables.md)
* <span data-ttu-id="ee6e6-123">[Intervalos](../../excel/excel-add-ins-ranges.md) e [APIs de Faixa Avançada](../../excel/excel-add-ins-ranges-advanced.md)</span><span class="sxs-lookup"><span data-stu-id="ee6e6-123">[Ranges](../../excel/excel-add-ins-ranges.md) and [Advanced Range APIs](../../excel/excel-add-ins-ranges-advanced.md)</span></span>
* [<span data-ttu-id="ee6e6-124">Formas</span><span class="sxs-lookup"><span data-stu-id="ee6e6-124">Shapes</span></span>](../../excel/excel-add-ins-shapes.md)
* [<span data-ttu-id="ee6e6-125">Tabelas</span><span class="sxs-lookup"><span data-stu-id="ee6e6-125">Tables</span></span>](../../excel/excel-add-ins-tables.md)
* [<span data-ttu-id="ee6e6-126">Pastas de trabalho e APIs no Nível do Aplicativo</span><span class="sxs-lookup"><span data-stu-id="ee6e6-126">Workbooks and Application-level APIs</span></span>](../../excel/excel-add-ins-workbooks.md)
* [<span data-ttu-id="ee6e6-127">Planilhas</span><span class="sxs-lookup"><span data-stu-id="ee6e6-127">Worksheets</span></span>](../../excel/excel-add-ins-worksheets.md)

<span data-ttu-id="ee6e6-128">Para saber mais sobre o modelo de objeto API JavaScript do Excel, consulte a [Documentação de referência da API JavaScript do Excel](/javascript/api/excel).</span><span class="sxs-lookup"><span data-stu-id="ee6e6-128">For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="ee6e6-129">Experimente amostras de código no Script Lab</span><span class="sxs-lookup"><span data-stu-id="ee6e6-129">Try out code samples in Script Lab</span></span>

<span data-ttu-id="ee6e6-130">Use o [Script Lab](../../overview/explore-with-script-lab.md) para começar a trabalhar rapidamente com um conjunto de exemplos internos que mostram como concluir tarefas com a API.</span><span class="sxs-lookup"><span data-stu-id="ee6e6-130">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="ee6e6-131">Você pode executar as amostras no Script Lab para ver instantaneamente o resultado no painel de tarefas ou planilha, examinar os exemplos para saber como a API funciona e até mesmo usar amostras para criar um protótipo do seu próprio suplemento.</span><span class="sxs-lookup"><span data-stu-id="ee6e6-131">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="ee6e6-132">Confira também</span><span class="sxs-lookup"><span data-stu-id="ee6e6-132">See also</span></span>

* [<span data-ttu-id="ee6e6-133">Documentação de Suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="ee6e6-133">Excel add-ins documentation</span></span>](../../excel/index.yml)
* [<span data-ttu-id="ee6e6-134">Visão geral dos suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="ee6e6-134">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
* [<span data-ttu-id="ee6e6-135">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="ee6e6-135">Excel JavaScript API reference</span></span>](/javascript/api/excel)
* [<span data-ttu-id="ee6e6-136">Disponibilidade de aplicativos e plataformas de cliente Office para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ee6e6-136">Office client application and platform availability for Office Add-ins</span></span>](../../overview/office-add-in-availability.md)
* [<span data-ttu-id="ee6e6-137">Usando o modelo de API específica do aplicativo</span><span class="sxs-lookup"><span data-stu-id="ee6e6-137">Using the application-specific API model</span></span>](../../develop/application-specific-api-model.md)
