---
title: Elemento AppDomain no arquivo de manifesto
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: da28b3b4dec5d669462a781db3c0628bd32c7182
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596785"
---
# <a name="appdomain-element"></a><span data-ttu-id="4fe98-102">Elemento AppDomain</span><span class="sxs-lookup"><span data-stu-id="4fe98-102">AppDomain element</span></span>

<span data-ttu-id="4fe98-103">Especifica domínios adicionais que carregam páginas na janela do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4fe98-103">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="4fe98-104">Ele também lista os domínios confiáveis dos quais as chamadas de API do Office. js podem ser feitas de IFrames no suplemento.</span><span class="sxs-lookup"><span data-stu-id="4fe98-104">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="4fe98-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="4fe98-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4fe98-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="4fe98-106">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="4fe98-107">O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="4fe98-107">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="4fe98-108">*Não* Coloque uma barra de fechamento, "/", no valor.</span><span class="sxs-lookup"><span data-stu-id="4fe98-108">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="4fe98-109">Contido em</span><span class="sxs-lookup"><span data-stu-id="4fe98-109">Contained in</span></span>

[<span data-ttu-id="4fe98-110">AppDomains</span><span class="sxs-lookup"><span data-stu-id="4fe98-110">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="4fe98-111">Comentários</span><span class="sxs-lookup"><span data-stu-id="4fe98-111">Remarks</span></span>

<span data-ttu-id="4fe98-112">Os elementos **AppDomain** deve ser usado para especificar os domínios adicionais diferentes daqueles especificados no elemento [SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="4fe98-112">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="4fe98-113">Confira mais informações em [Manifesto XML de Suplementos do Office](../../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="4fe98-113">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
