---
title: Elemento AppDomain no arquivo de manifesto
description: ''
ms.date: 05/15/2019
localization_priority: Normal
ms.openlocfilehash: b1d71648cc7646eec246f3d0a8113c843eed2e74
ms.sourcegitcommit: b0e71ae0ae09c57b843d4de277081845c108a645
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2019
ms.locfileid: "34337192"
---
# <a name="appdomain-element"></a><span data-ttu-id="31645-102">Elemento AppDomain</span><span class="sxs-lookup"><span data-stu-id="31645-102">AppDomain element</span></span>

<span data-ttu-id="31645-103">Especifica um domínio adicional que será usado para carregar páginas na janela do suplemento.</span><span class="sxs-lookup"><span data-stu-id="31645-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="31645-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="31645-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="31645-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="31645-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="31645-106">O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="31645-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="31645-107">*Não* Coloque uma barra de fechamento, "/", no valor.</span><span class="sxs-lookup"><span data-stu-id="31645-107">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="31645-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="31645-108">Contained in</span></span>

[<span data-ttu-id="31645-109">AppDomains</span><span class="sxs-lookup"><span data-stu-id="31645-109">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="31645-110">Comentários</span><span class="sxs-lookup"><span data-stu-id="31645-110">Remarks</span></span>

<span data-ttu-id="31645-111">Os elementos **AppDomain** deve ser usado para especificar os domínios adicionais diferentes daqueles especificados no elemento [SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="31645-111">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="31645-112">Confira mais informações em [Manifesto XML de Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="31645-112">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
