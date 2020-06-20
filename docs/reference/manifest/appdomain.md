---
title: Elemento AppDomain no arquivo de manifesto
description: Especifica domínios adicionais que são usados pelo seu suplemento e que deve ser confiável para o Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: ae49944afceada559b39353cd119e26a21fd3d15
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778645"
---
# <a name="appdomain-element"></a><span data-ttu-id="cc6a9-103">Elemento AppDomain</span><span class="sxs-lookup"><span data-stu-id="cc6a9-103">AppDomain element</span></span>

<span data-ttu-id="cc6a9-104">Especifica um domínio adicional no qual o Office deve confiar, além do especificado no [elemento SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="cc6a9-104">Specifies an additional domain that Office should trust, in addition to the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="cc6a9-105">A especificação de um domínio tem estes efeitos:</span><span class="sxs-lookup"><span data-stu-id="cc6a9-105">Specifying a domain has these effects:</span></span>

- <span data-ttu-id="cc6a9-106">Ele permite que páginas, rotas ou outros recursos no domínio sejam abertos diretamente no painel de tarefas raiz do suplemento em plataformas do Office.</span><span class="sxs-lookup"><span data-stu-id="cc6a9-106">It enables pages, routes, or other resources in the domain to be opened directly in the root task pane of the add-in on desktop Office platforms.</span></span> <span data-ttu-id="cc6a9-107">(Especificar um domínio em um **AppDomain** não é necessário para o Office na Web ou para abrir um recurso em um iframe, nem é necessário para abrir um recurso em uma caixa de diálogo aberta com a [API da caixa de diálogo](../../develop/dialog-api-in-office-add-ins.md).)</span><span class="sxs-lookup"><span data-stu-id="cc6a9-107">(Specifying a domain in an **AppDomain** isn't necessary for Office on the web or to open a resource in an IFrame, nor it is necessary for opening a resource in a dialog opened with the [Dialog API](../../develop/dialog-api-in-office-add-ins.md).)</span></span>
- <span data-ttu-id="cc6a9-108">Ele permite que as páginas no domínio façam chamadas de API Office.js de IFrames no suplemento.</span><span class="sxs-lookup"><span data-stu-id="cc6a9-108">It enables pages in the domain to make Office.js API calls from IFrames within the add-in.</span></span>

<span data-ttu-id="cc6a9-109">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="cc6a9-109">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="cc6a9-110">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="cc6a9-110">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="cc6a9-111">O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain.com</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="cc6a9-111">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain.com</AppDomain>`).</span></span>
> 2. <span data-ttu-id="cc6a9-112">Se houver uma porta explícita para o domínio, inclua-a (por exemplo, `<AppDomain>https://myappdomain.com:9999</AppDomain>` ).</span><span class="sxs-lookup"><span data-stu-id="cc6a9-112">If there is an explicit port for the domain, include it (e.g.,`<AppDomain>https://myappdomain.com:9999</AppDomain>`).</span></span>
> 3. <span data-ttu-id="cc6a9-113">Se um subdomínio precisar ser confiável, inclua-o (por exemplo, `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` ).</span><span class="sxs-lookup"><span data-stu-id="cc6a9-113">If a subdomain needs to be trusted, include it (e.g.,`<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>`).</span></span> <span data-ttu-id="cc6a9-114">O subdomínio `mysubdomain.mydomain.com` e os `mydomain.com` domínios são diferentes.</span><span class="sxs-lookup"><span data-stu-id="cc6a9-114">The subdomain `mysubdomain.mydomain.com` and `mydomain.com` are different domains.</span></span> <span data-ttu-id="cc6a9-115">Se ambos precisam ser confiáveis, então ambos precisam estar em elementos **AppDomain** separados.</span><span class="sxs-lookup"><span data-stu-id="cc6a9-115">If both need to be trusted, then both need to be in separate **AppDomain** elements.</span></span>
> 4. <span data-ttu-id="cc6a9-116">A listagem do mesmo domínio que o especificado no [elemento SourceLocation](sourcelocation.md) não tem efeito e pode ser enganosa.</span><span class="sxs-lookup"><span data-stu-id="cc6a9-116">Listing the same domain as the one specified in the [SourceLocation element](sourcelocation.md) has no effect and may be misleading.</span></span> <span data-ttu-id="cc6a9-117">Em particular, quando você está desenvolvendo `localhost` , não é necessário criar um elemento **AppDomain** para `localhost` .</span><span class="sxs-lookup"><span data-stu-id="cc6a9-117">In particular, when you are developing on `localhost`, you don't need to create an **AppDomain** element for `localhost`.</span></span>
> 5. <span data-ttu-id="cc6a9-118">Não inclua nenhum segmento de uma URL além do domínio.</span><span class="sxs-lookup"><span data-stu-id="cc6a9-118">Don't include any segments of a URL past the domain.</span></span> <span data-ttu-id="cc6a9-119">Por exemplo, não inclua a URL completa de uma página.</span><span class="sxs-lookup"><span data-stu-id="cc6a9-119">For example, don't include the full URL of a page.</span></span>
> 6. <span data-ttu-id="cc6a9-120">*Não* Coloque uma barra de fechamento, "/", no valor.</span><span class="sxs-lookup"><span data-stu-id="cc6a9-120">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="cc6a9-121">Contido em</span><span class="sxs-lookup"><span data-stu-id="cc6a9-121">Contained in</span></span>

[<span data-ttu-id="cc6a9-122">AppDomains</span><span class="sxs-lookup"><span data-stu-id="cc6a9-122">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="cc6a9-123">Comentários</span><span class="sxs-lookup"><span data-stu-id="cc6a9-123">Remarks</span></span>

<span data-ttu-id="cc6a9-124">Para saber mais, confira o [manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="cc6a9-124">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
