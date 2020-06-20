---
title: Elemento AppDomains no arquivo de manifesto
description: Lista todos os domínios, além do domínio especificado no `SourceLocation` elemento que seu suplemento do Office usará e deve ser confiável para o Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 751e4ad2ffa5fd50739a855fad48964473b154f1
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778652"
---
# <a name="appdomains-element"></a><span data-ttu-id="60f16-103">Elemento AppDomains</span><span class="sxs-lookup"><span data-stu-id="60f16-103">AppDomains element</span></span>

<span data-ttu-id="60f16-104">Lista todos os domínios, além do domínio especificado no `SourceLocation` elemento, que o seu suplemento do Office usará e que deve ser confiável para o Office.</span><span class="sxs-lookup"><span data-stu-id="60f16-104">Lists any domains, in addition to the domain specified in the `SourceLocation` element, that your Office Add-in will use and that should be trusted by Office.</span></span> <span data-ttu-id="60f16-105">Isso permite que as páginas nos domínios façam chamadas para Office.js APIs de IFrames no suplemento e têm outros efeitos.</span><span class="sxs-lookup"><span data-stu-id="60f16-105">This enables pages in the domains to make calls to Office.js APIs from IFrames within the add-in and has other effects.</span></span> <span data-ttu-id="60f16-106">Para cada domínio adicional, especifique um elemento **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="60f16-106">For each additional domain, specify an **AppDomain** element.</span></span>

 <span data-ttu-id="60f16-107">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="60f16-107">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="60f16-108">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="60f16-108">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="60f16-109">Há restrições sobre o que pode ser o valor de um elemento **AppDomain** .</span><span class="sxs-lookup"><span data-stu-id="60f16-109">There are restrictions on what can be the value of a **AppDomain** element.</span></span> <span data-ttu-id="60f16-110">Para obter mais informações, consulte [AppDomain](appdomain.md).</span><span class="sxs-lookup"><span data-stu-id="60f16-110">For more information, see [AppDomain](appdomain.md).</span></span>

## <a name="contained-in"></a><span data-ttu-id="60f16-111">Contido em</span><span class="sxs-lookup"><span data-stu-id="60f16-111">Contained in</span></span>

[<span data-ttu-id="60f16-112">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="60f16-112">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="60f16-113">Pode conter</span><span class="sxs-lookup"><span data-stu-id="60f16-113">Can contain</span></span>

[<span data-ttu-id="60f16-114">AppDomain</span><span class="sxs-lookup"><span data-stu-id="60f16-114">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="60f16-115">Comentários</span><span class="sxs-lookup"><span data-stu-id="60f16-115">Remarks</span></span>

<span data-ttu-id="60f16-116">Por padrão, o seu suplemento pode carregar qualquer página que esteja no mesmo domínio que o local especificado no elemento [SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="60f16-116">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="60f16-117">Esse elemento não pode estar vazio.</span><span class="sxs-lookup"><span data-stu-id="60f16-117">This element can't be empty.</span></span>
