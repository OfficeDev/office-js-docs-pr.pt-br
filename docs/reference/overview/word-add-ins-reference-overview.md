---
title: Visão geral da API JavaScript do Word
description: Visão geral da API JavaScript do Word.
ms.date: 02/19/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: dfc87a8f9f5c7048262d9c2889ae68eb38c0ed30
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719886"
---
# <a name="word-javascript-api-overview"></a><span data-ttu-id="8a9f6-103">Visão geral da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="8a9f6-103">Word JavaScript API overview</span></span>

<span data-ttu-id="8a9f6-104">Um suplemento do Word interage com objetos no Word usando a API JavaScript do Office, que inclui dois modelos de objetos JavaScript:</span><span class="sxs-lookup"><span data-stu-id="8a9f6-104">An Word add-in interacts with objects in Word by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="8a9f6-105">**API JavaScript do Word**: introduzida com o Office 2016, a [API JavaScript do Word](/javascript/api/word) fornece objetos fortemente tipados que você pode usar para acessar objetos e metadados em um documento do Word.</span><span class="sxs-lookup"><span data-stu-id="8a9f6-105">**Word JavaScript API**: Introduced with Office 2016, the [Word JavaScript API](/javascript/api/word) provides strongly-typed objects that you can use to access objects and metadata in a Word document.</span></span> 

* <span data-ttu-id="8a9f6-106">**APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="8a9f6-106">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="8a9f6-107">Esta seção da documentação concentra-se na API JavaScript do Word, que você usará para desenvolver a maior parte da funcionalidade em suplementos direcionados para o Word na Web ou para o Word 2016 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="8a9f6-107">This section of the documentation focuses on the Word JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Word on the web or Word 2016 or later.</span></span> <span data-ttu-id="8a9f6-108">Para saber mais sobre a API Comum, confira [Modelo de objeto da API JavaScript comum](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="8a9f6-108">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="8a9f6-109">Aprenda conceitos de programação</span><span class="sxs-lookup"><span data-stu-id="8a9f6-109">Learn programming concepts</span></span>

<span data-ttu-id="8a9f6-110">Veja [Conceitos fundamentais de programação com a API JavaScript do Word](../../word/word-add-ins-core-concepts.md) para obter informações sobre conceitos de programação importantes.</span><span class="sxs-lookup"><span data-stu-id="8a9f6-110">See [Fundamental programming concepts with the Word JavaScript API](../../word/word-add-ins-core-concepts.md) for information about important programming concepts.</span></span>
 
## <a name="learn-about-api-capabilities"></a><span data-ttu-id="8a9f6-111">Saiba mais sobre recursos da API</span><span class="sxs-lookup"><span data-stu-id="8a9f6-111">Learn about API capabilities</span></span>

<span data-ttu-id="8a9f6-112">Use outros artigos nesta seção da documentação para saber [como obter o documento inteiro de um suplemento](../../word/get-the-whole-document-from-an-add-in-for-word.md), [usar as opções de pesquisa para localizar o texto no suplemento do Word](../../word/search-option-guidance.md) e muito mais.</span><span class="sxs-lookup"><span data-stu-id="8a9f6-112">Use other articles in this section of the documentation to learn how to [get the whole document from an add-in](../../word/get-the-whole-document-from-an-add-in-for-word.md), [use search options to find text in your Word add-in](../../word/search-option-guidance.md), and more.</span></span> <span data-ttu-id="8a9f6-113">Confira o Sumário para obter a lista completa de artigos disponíveis.</span><span class="sxs-lookup"><span data-stu-id="8a9f6-113">See the table of contents for the complete list of available articles.</span></span>

<span data-ttu-id="8a9f6-114">Para ter a experiência prática com o uso da API JavaScript do Word para acessar objetos no Word, conclua o [tutorial do suplemento do Word](../../tutorials/word-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="8a9f6-114">For hands-on experience using the Word JavaScript API to access objects in Word, complete the [Word add-in tutorial](../../tutorials/word-tutorial.md).</span></span> 

<span data-ttu-id="8a9f6-115">Para saber mais sobre o modelo de objeto API JavaScript do Word, consulte a [Documentação de referência da API JavaScript do Word](/javascript/api/word).</span><span class="sxs-lookup"><span data-stu-id="8a9f6-115">For detailed information about the Word JavaScript API object model, see the [Word JavaScript API reference documentation](/javascript/api/word).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="8a9f6-116">Experimente amostras de código no Script Lab</span><span class="sxs-lookup"><span data-stu-id="8a9f6-116">Try out code samples in Script Lab</span></span>

<span data-ttu-id="8a9f6-117">Use o [Script Lab](../../overview/explore-with-script-lab.md) para começar a trabalhar rapidamente com um conjunto de exemplos internos que mostram como concluir tarefas com a API.</span><span class="sxs-lookup"><span data-stu-id="8a9f6-117">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="8a9f6-118">Você pode executar as amostras no Script Lab para ver instantaneamente o resultado no painel de tarefas ou documento, examinar os exemplos para saber como a API funciona e até mesmo usar amostras para criar um protótipo do seu próprio suplemento.</span><span class="sxs-lookup"><span data-stu-id="8a9f6-118">You can run the samples in Script Lab to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="8a9f6-119">Confira também</span><span class="sxs-lookup"><span data-stu-id="8a9f6-119">See also</span></span>

- [<span data-ttu-id="8a9f6-120">Documentação de suplementos do Word</span><span class="sxs-lookup"><span data-stu-id="8a9f6-120">Word add-ins documentation</span></span>](../../word/index.md)
- [<span data-ttu-id="8a9f6-121">Visão geral dos suplementos do Word</span><span class="sxs-lookup"><span data-stu-id="8a9f6-121">Word add-ins overview</span></span>](../../word/word-add-ins-programming-overview.md)
- [<span data-ttu-id="8a9f6-122">Referências da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="8a9f6-122">Word JavaScript API reference</span></span>](/javascript/api/word)
- [<span data-ttu-id="8a9f6-123">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8a9f6-123">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
- [<span data-ttu-id="8a9f6-124">Especificações abertas da API</span><span class="sxs-lookup"><span data-stu-id="8a9f6-124">API open specifications</span></span>](../openspec/openspec.md)
