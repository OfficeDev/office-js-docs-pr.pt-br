---
title: Visão geral da API JavaScript do OneNote
description: Saiba mais sobre a API JavaScript do OneNote
ms.date: 07/28/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 08e98e81e46ca62178235454d3ba44f35be2eec2
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293629"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="b0fef-103">Visão geral da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="b0fef-103">OneNote JavaScript API overview</span></span>

<span data-ttu-id="b0fef-104">Um suplemento do OneNote interage com objetos no OneNote na Web usando a API JavaScript do Office, que inclui dois modelos de objetos JavaScript:</span><span class="sxs-lookup"><span data-stu-id="b0fef-104">A OneNote add-in interacts with objects in OneNote on the web by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="b0fef-105">**API de JavaScript do Excel**: estas são as [APIs específicas do aplicativo](../../develop/application-specific-api-model.md) para o Excel.</span><span class="sxs-lookup"><span data-stu-id="b0fef-105">**OneNote JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for OneNote.</span></span> <span data-ttu-id="b0fef-106">Introduzida com o Office 2016, a [API de JavaScript do OneNote](/javascript/api/onenote) fornece objetos de tipo forte que você pode usar para acessar objetos no OneNote na Web.</span><span class="sxs-lookup"><span data-stu-id="b0fef-106">Introduced with Office 2016, the [OneNote JavaScript API](/javascript/api/onenote) provides strongly-typed objects that you can use to access objects in OneNote on the web.</span></span>

* <span data-ttu-id="b0fef-107">**APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="b0fef-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="b0fef-108">Esta seção da documentação concentra-se na API JavaScript do OneNote, que você usará para desenvolver a maior parte da funcionalidade em suplementos direcionados para o OneNote na Web.</span><span class="sxs-lookup"><span data-stu-id="b0fef-108">This section of the documentation focuses on the OneNote JavaScript API, which you'll use to develop the majority of functionality in add-ins that target OneNote on the web.</span></span> <span data-ttu-id="b0fef-109">Para obter mais informações do API comum, consulte [Modelo do objeto do JavaScript API comum](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="b0fef-109">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span>

## <a name="learn-programming-concepts"></a><span data-ttu-id="b0fef-110">Aprenda conceitos de programação</span><span class="sxs-lookup"><span data-stu-id="b0fef-110">Learn programming concepts</span></span>

<span data-ttu-id="b0fef-111">Confira os artigos a seguir para obter informações sobre conceitos de programação importantes:</span><span class="sxs-lookup"><span data-stu-id="b0fef-111">See the following articles for information about important programming concepts:</span></span>

* [<span data-ttu-id="b0fef-112">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="b0fef-112">OneNote JavaScript API programming overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
* [<span data-ttu-id="b0fef-113">Trabalhar com conteúdo da página do OneNote</span><span class="sxs-lookup"><span data-stu-id="b0fef-113">Work with OneNote page content</span></span>](../../onenote/onenote-add-ins-page-content.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="b0fef-114">Saiba mais sobre recursos da API</span><span class="sxs-lookup"><span data-stu-id="b0fef-114">Learn about API capabilities</span></span>

<span data-ttu-id="b0fef-115">Para ter experiência prática com o uso do API JavaScript do OneNote para interagir com o conteúdo no OneNote na Web, preencha o [início rápido do suplemento do OneNote](../../quickstarts/onenote-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="b0fef-115">For hands-on experience using the OneNote JavaScript API to interact with content in OneNote on the web, complete the [OneNote add-in quick start](../../quickstarts/onenote-quickstart.md).</span></span>

<span data-ttu-id="b0fef-116">Para saber mais sobre a API JavaScript do OneNote, consulte a [documentação de referência da API JavaScript do OneNote](/javascript/api/onenote).</span><span class="sxs-lookup"><span data-stu-id="b0fef-116">For detailed information about the OneNote JavaScript API object model, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="b0fef-117">Confira também</span><span class="sxs-lookup"><span data-stu-id="b0fef-117">See also</span></span>

* [<span data-ttu-id="b0fef-118">Documentação de Suplementos do OneNote</span><span class="sxs-lookup"><span data-stu-id="b0fef-118">OneNote add-ins documentation</span></span>](../../onenote/index.yml)
* [<span data-ttu-id="b0fef-119">Visão geral dos suplementos do OneNote</span><span class="sxs-lookup"><span data-stu-id="b0fef-119">OneNote add-ins overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
* [<span data-ttu-id="b0fef-120">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="b0fef-120">OneNote JavaScript API reference</span></span>](/javascript/api/onenote)
* [<span data-ttu-id="b0fef-121">Disponibilidade de aplicativos e plataformas de cliente Office para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b0fef-121">Office client application and platform availability for Office Add-ins</span></span>](../../overview/office-add-in-availability.md)
