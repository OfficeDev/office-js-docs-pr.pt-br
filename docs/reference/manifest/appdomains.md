---
title: Elemento AppDomains no arquivo de manifesto
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: b6db3d46d004021f25edd5733566544010abb457
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575328"
---
# <a name="appdomains-element"></a><span data-ttu-id="1bce7-102">Elemento AppDomains</span><span class="sxs-lookup"><span data-stu-id="1bce7-102">AppDomains element</span></span>

<span data-ttu-id="1bce7-103">Lista todos os domínios, além do domínio especificado no `SourceLocation` elemento que seu suplemento do Office usará para carregar páginas.</span><span class="sxs-lookup"><span data-stu-id="1bce7-103">Lists any domains in addition to the domain specified in the `SourceLocation` element that your Office Add-in will use to load pages.</span></span> <span data-ttu-id="1bce7-104">Ele também lista os domínios confiáveis dos quais as chamadas de API do Office. js podem ser feitas de IFrames no suplemento.</span><span class="sxs-lookup"><span data-stu-id="1bce7-104">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span> <span data-ttu-id="1bce7-105">Para cada domínio adicional, especifique um elemento AppDomain.</span><span class="sxs-lookup"><span data-stu-id="1bce7-105">For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="1bce7-106">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="1bce7-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1bce7-107">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="1bce7-107">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="1bce7-108">O valor de cada elemento **AppDomain** deve incluir o protocolo (por exemplo, `<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="1bce7-108">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="1bce7-109">Contido em</span><span class="sxs-lookup"><span data-stu-id="1bce7-109">Contained in</span></span>

[<span data-ttu-id="1bce7-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="1bce7-110">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="1bce7-111">Pode conter</span><span class="sxs-lookup"><span data-stu-id="1bce7-111">Can contain</span></span>

[<span data-ttu-id="1bce7-112">AppDomain</span><span class="sxs-lookup"><span data-stu-id="1bce7-112">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="1bce7-113">Comentários</span><span class="sxs-lookup"><span data-stu-id="1bce7-113">Remarks</span></span>

<span data-ttu-id="1bce7-114">Por padrão, o seu suplemento pode carregar qualquer página que esteja no mesmo domínio que o local especificado no elemento [SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="1bce7-114">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="1bce7-115">Para carregar páginas que não estejam no mesmo domínio do que o suplemento, especifique os domínios usando os elementos **AppDomains** e **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="1bce7-115">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="1bce7-116">Esse elemento não pode estar vazio.</span><span class="sxs-lookup"><span data-stu-id="1bce7-116">This element can't be empty.</span></span>
