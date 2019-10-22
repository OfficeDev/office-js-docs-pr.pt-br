---
title: Elemento AppDomain no arquivo de manifesto
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 2f65302d1ac3d85f2867cd13501bc67606cd00b5
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/21/2019
ms.locfileid: "35575636"
---
# <a name="appdomain-element"></a><span data-ttu-id="4de54-102">Elemento AppDomain</span><span class="sxs-lookup"><span data-stu-id="4de54-102">AppDomain element</span></span>

<span data-ttu-id="4de54-103">Especifica domínios adicionais que carregam páginas na janela do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4de54-103">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="4de54-104">Ele também lista os domínios confiáveis dos quais as chamadas de API do Office. js podem ser feitas de IFrames no suplemento.</span><span class="sxs-lookup"><span data-stu-id="4de54-104">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="4de54-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="4de54-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4de54-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="4de54-106">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="4de54-107">O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="4de54-107">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="4de54-108">*Não* Coloque uma barra de fechamento, "/", no valor.</span><span class="sxs-lookup"><span data-stu-id="4de54-108">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="4de54-109">Contido em</span><span class="sxs-lookup"><span data-stu-id="4de54-109">Contained in</span></span>

[<span data-ttu-id="4de54-110">AppDomains</span><span class="sxs-lookup"><span data-stu-id="4de54-110">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="4de54-111">Comentários</span><span class="sxs-lookup"><span data-stu-id="4de54-111">Remarks</span></span>

<span data-ttu-id="4de54-112">Os elementos **AppDomain** deve ser usado para especificar os domínios adicionais diferentes daqueles especificados no elemento [SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="4de54-112">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="4de54-113">Confira mais informações em [Manifesto XML de Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="4de54-113">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
