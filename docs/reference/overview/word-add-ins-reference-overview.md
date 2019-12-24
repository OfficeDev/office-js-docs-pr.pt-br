---
title: Visão geral da API JavaScript do Word
description: ''
ms.date: 07/05/2019
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 6728c7491d84f2bc044f7e5a3199ad6d90979628
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851254"
---
# <a name="word-javascript-api-overview"></a><span data-ttu-id="c2358-102">Visão geral da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="c2358-102">Word JavaScript API overview</span></span>

<span data-ttu-id="c2358-103">Um suplemento do Word interage com objetos no Word usando a API JavaScript para Office, que inclui dois modelos de objeto JavaScript:</span><span class="sxs-lookup"><span data-stu-id="c2358-103">An Word add-in interacts with objects in Word by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="c2358-104">**API JavaScript do Word**: introduzida com o Office 2016, a [API JavaScript do Word](/javascript/api/word) fornece objetos fortemente tipados que você pode usar para acessar objetos e metadados em um documento do Word.</span><span class="sxs-lookup"><span data-stu-id="c2358-104">**Word JavaScript API**: Introduced with Office 2016, the [Word JavaScript API](/javascript/api/word) provides strongly-typed objects that you can use to access objects and metadata in a Word document.</span></span> 

* <span data-ttu-id="c2358-105">**APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="c2358-105">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="c2358-106">Esta seção da documentação concentra-se na API JavaScript do Word, que você usará para desenvolver a maior parte da funcionalidade em suplementos direcionados para o Word na Web ou para o Word 2016 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="c2358-106">This section of the documentation focuses on the Word JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Word on the web or Word 2016 or later.</span></span> <span data-ttu-id="c2358-107">Para saber mais sobre a API comum, confira [ modelo de objeto de API do JavaScript para Office](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="c2358-107">For information about the Common API, see [JavaScript API for Office](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="c2358-108">Aprenda conceitos de programação</span><span class="sxs-lookup"><span data-stu-id="c2358-108">Learn programming concepts</span></span>

<span data-ttu-id="c2358-109">Veja [Conceitos fundamentais de programação com a API JavaScript do Word](../../word/word-add-ins-core-concepts.md) para obter informações sobre conceitos de programação importantes.</span><span class="sxs-lookup"><span data-stu-id="c2358-109">See [Fundamental programming concepts with the Word JavaScript API](../../word/word-add-ins-core-concepts.md) for information about important programming concepts.</span></span>
 
## <a name="learn-about-api-capabilities"></a><span data-ttu-id="c2358-110">Saiba mais sobre recursos da API</span><span class="sxs-lookup"><span data-stu-id="c2358-110">Learn about API capabilities</span></span>

<span data-ttu-id="c2358-111">Use outros artigos nesta seção da documentação para saber [como obter o documento inteiro de um suplemento](../../word/get-the-whole-document-from-an-add-in-for-word.md), [usar as opções de pesquisa para localizar o texto no suplemento do Word](../../word/search-option-guidance.md) e muito mais.</span><span class="sxs-lookup"><span data-stu-id="c2358-111">Use other articles in this section of the documentation to learn how to [get the whole document from an add-in](../../word/get-the-whole-document-from-an-add-in-for-word.md), [use search options to find text in your Word add-in](../../word/search-option-guidance.md), and more.</span></span> <span data-ttu-id="c2358-112">Confira o Sumário para obter a lista completa de artigos disponíveis.</span><span class="sxs-lookup"><span data-stu-id="c2358-112">See the table of contents for the complete list of available articles.</span></span>

<span data-ttu-id="c2358-113">Para ter a experiência prática com o uso da API JavaScript do Word para acessar objetos no Word, conclua o [tutorial do suplemento do Word](../../tutorials/word-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="c2358-113">For hands-on experience using the Word JavaScript API to access objects in Word, complete the [Word add-in tutorial](../../tutorials/word-tutorial.md).</span></span> 

<span data-ttu-id="c2358-114">Para saber mais sobre o modelo de objeto API JavaScript do Word, consulte a [Documentação de referência da API JavaScript do Word](/javascript/api/word).</span><span class="sxs-lookup"><span data-stu-id="c2358-114">For detailed information about the Word JavaScript API object model, see the [Word JavaScript API reference documentation](/javascript/api/word).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="c2358-115">Experimente amostras de código no Script Lab</span><span class="sxs-lookup"><span data-stu-id="c2358-115">Try out code samples in Script Lab</span></span>

<span data-ttu-id="c2358-116">Use o [Script Lab](../../overview/explore-with-script-lab.md) para começar a trabalhar rapidamente com um conjunto de exemplos internos que mostram como concluir tarefas com a API.</span><span class="sxs-lookup"><span data-stu-id="c2358-116">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="c2358-117">Você pode executar as amostras no Script Lab para ver instantaneamente o resultado no painel de tarefas ou documento, examinar os exemplos para saber como a API funciona e até mesmo usar amostras para criar um protótipo do seu próprio suplemento.</span><span class="sxs-lookup"><span data-stu-id="c2358-117">You can run the samples in Script Lab to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="c2358-118">Confira também</span><span class="sxs-lookup"><span data-stu-id="c2358-118">See also</span></span>

- [<span data-ttu-id="c2358-119">Documentação de suplementos do Word</span><span class="sxs-lookup"><span data-stu-id="c2358-119">Word add-ins documentation</span></span>](../../word/index.md)
- [<span data-ttu-id="c2358-120">Visão geral dos suplementos do Word</span><span class="sxs-lookup"><span data-stu-id="c2358-120">Word add-ins overview</span></span>](../../word/word-add-ins-programming-overview.md)
- [<span data-ttu-id="c2358-121">Referências da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="c2358-121">Word JavaScript API reference</span></span>](/javascript/api/word)
- [<span data-ttu-id="c2358-122">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="c2358-122">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
- [<span data-ttu-id="c2358-123">Especificações abertas da API</span><span class="sxs-lookup"><span data-stu-id="c2358-123">API open specifications</span></span>](../openspec/openspec.md)
