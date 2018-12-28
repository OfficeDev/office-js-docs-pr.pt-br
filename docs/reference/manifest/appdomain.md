---
title: Elemento AppDomain no arquivo de manifesto
description: ''
ms.date: 12/13/2018
ms.openlocfilehash: 2b55f2c1ea7a2a3dc7dec42c913d74006c0f2e3b
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433065"
---
# <a name="appdomain-element"></a><span data-ttu-id="033da-102">Elemento AppDomain</span><span class="sxs-lookup"><span data-stu-id="033da-102">AppDomain element</span></span>

<span data-ttu-id="033da-103">Especifica um domínio adicional que será usado para carregar páginas na janela do suplemento.</span><span class="sxs-lookup"><span data-stu-id="033da-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="033da-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="033da-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="033da-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="033da-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> <span data-ttu-id="033da-106">O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="033da-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="033da-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="033da-107">Contained in</span></span>

[<span data-ttu-id="033da-108">AppDomains</span><span class="sxs-lookup"><span data-stu-id="033da-108">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="033da-109">Comentários</span><span class="sxs-lookup"><span data-stu-id="033da-109">Remarks</span></span>

<span data-ttu-id="033da-110">Os elementos **AppDomain** deve ser usado para especificar os domínios adicionais diferentes daqueles especificados no elemento [SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="033da-110">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="033da-111">Confira mais informações em [Manifesto XML de Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="033da-111">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
