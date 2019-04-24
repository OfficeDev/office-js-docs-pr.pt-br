---
title: Elemento AppDomain no arquivo de manifesto
description: ''
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: 8216603c87a7dcafde84d25a82f068c9aa86ed96
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450748"
---
# <a name="appdomain-element"></a><span data-ttu-id="59638-102">Elemento AppDomain</span><span class="sxs-lookup"><span data-stu-id="59638-102">AppDomain element</span></span>

<span data-ttu-id="59638-103">Especifica um domínio adicional que será usado para carregar páginas na janela do suplemento.</span><span class="sxs-lookup"><span data-stu-id="59638-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="59638-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="59638-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="59638-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="59638-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="59638-106">O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="59638-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="59638-107">*Não* Coloque uma barra de fechamento, "/", no valor.</span><span class="sxs-lookup"><span data-stu-id="59638-107">Do *not* put a closing slash, "/", on the the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="59638-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="59638-108">Contained in</span></span>

[<span data-ttu-id="59638-109">AppDomains</span><span class="sxs-lookup"><span data-stu-id="59638-109">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="59638-110">Comentários</span><span class="sxs-lookup"><span data-stu-id="59638-110">Remarks</span></span>

<span data-ttu-id="59638-111">Os elementos **AppDomain** deve ser usado para especificar os domínios adicionais diferentes daqueles especificados no elemento [SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="59638-111">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="59638-112">Confira mais informações em [Manifesto XML de Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="59638-112">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
