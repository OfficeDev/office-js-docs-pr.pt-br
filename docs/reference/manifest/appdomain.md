---
title: Elemento AppDomain no arquivo de manifesto
description: Especifica domínios adicionais que carregam páginas na janela do suplemento.
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: ddacae6d8aa45ccccd3a8acbb42de48b152fb9d2
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608772"
---
# <a name="appdomain-element"></a><span data-ttu-id="4e071-103">Elemento AppDomain</span><span class="sxs-lookup"><span data-stu-id="4e071-103">AppDomain element</span></span>

<span data-ttu-id="4e071-104">Especifica domínios adicionais que carregam páginas na janela do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4e071-104">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="4e071-105">Ele também lista os domínios confiáveis dos quais as chamadas de API do Office. js podem ser feitas de IFrames no suplemento.</span><span class="sxs-lookup"><span data-stu-id="4e071-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="4e071-106">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="4e071-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4e071-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="4e071-107">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="4e071-108">O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="4e071-108">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="4e071-109">*Não* Coloque uma barra de fechamento, "/", no valor.</span><span class="sxs-lookup"><span data-stu-id="4e071-109">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="4e071-110">Contido em</span><span class="sxs-lookup"><span data-stu-id="4e071-110">Contained in</span></span>

[<span data-ttu-id="4e071-111">AppDomains</span><span class="sxs-lookup"><span data-stu-id="4e071-111">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="4e071-112">Comentários</span><span class="sxs-lookup"><span data-stu-id="4e071-112">Remarks</span></span>

<span data-ttu-id="4e071-113">Os elementos **AppDomain** deve ser usado para especificar os domínios adicionais diferentes daqueles especificados no elemento [SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="4e071-113">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="4e071-114">Confira mais informações em [Manifesto XML de Suplementos do Office](../../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="4e071-114">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
