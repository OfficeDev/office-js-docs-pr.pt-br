---
title: Elemento AppDomain no arquivo de manifesto
description: Especifica domínios adicionais que carregam páginas na janela do suplemento.
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 6990f759df806f24b1d617c036bc1a452e6da38f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718451"
---
# <a name="appdomain-element"></a><span data-ttu-id="4932d-103">Elemento AppDomain</span><span class="sxs-lookup"><span data-stu-id="4932d-103">AppDomain element</span></span>

<span data-ttu-id="4932d-104">Especifica domínios adicionais que carregam páginas na janela do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4932d-104">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="4932d-105">Ele também lista os domínios confiáveis dos quais as chamadas de API do Office. js podem ser feitas de IFrames no suplemento.</span><span class="sxs-lookup"><span data-stu-id="4932d-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="4932d-106">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="4932d-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4932d-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="4932d-107">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="4932d-108">O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="4932d-108">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="4932d-109">*Não* Coloque uma barra de fechamento, "/", no valor.</span><span class="sxs-lookup"><span data-stu-id="4932d-109">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="4932d-110">Contido em</span><span class="sxs-lookup"><span data-stu-id="4932d-110">Contained in</span></span>

[<span data-ttu-id="4932d-111">AppDomains</span><span class="sxs-lookup"><span data-stu-id="4932d-111">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="4932d-112">Comentários</span><span class="sxs-lookup"><span data-stu-id="4932d-112">Remarks</span></span>

<span data-ttu-id="4932d-113">Os elementos **AppDomain** deve ser usado para especificar os domínios adicionais diferentes daqueles especificados no elemento [SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="4932d-113">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="4932d-114">Confira mais informações em [Manifesto XML de Suplementos do Office](../../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="4932d-114">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
