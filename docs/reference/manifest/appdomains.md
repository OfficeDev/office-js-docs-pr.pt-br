---
title: Elemento AppDomains no arquivo de manifesto
description: Lista todos os domínios, além do domínio especificado no `SourceLocation` elemento que seu suplemento do Office usará para carregar páginas.
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 9183f1815e97bd8d4ac1a7e2cf72d5547d153f7e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608765"
---
# <a name="appdomains-element"></a><span data-ttu-id="51a3a-103">Elemento AppDomains</span><span class="sxs-lookup"><span data-stu-id="51a3a-103">AppDomains element</span></span>

<span data-ttu-id="51a3a-104">Lista todos os domínios, além do domínio especificado no `SourceLocation` elemento que seu suplemento do Office usará para carregar páginas.</span><span class="sxs-lookup"><span data-stu-id="51a3a-104">Lists any domains in addition to the domain specified in the `SourceLocation` element that your Office Add-in will use to load pages.</span></span> <span data-ttu-id="51a3a-105">Ele também lista os domínios confiáveis dos quais as chamadas de API do Office. js podem ser feitas de IFrames no suplemento.</span><span class="sxs-lookup"><span data-stu-id="51a3a-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span> <span data-ttu-id="51a3a-106">Para cada domínio adicional, especifique um elemento AppDomain.</span><span class="sxs-lookup"><span data-stu-id="51a3a-106">For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="51a3a-107">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="51a3a-107">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="51a3a-108">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="51a3a-108">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="51a3a-109">O valor de cada elemento **AppDomain** deve incluir o protocolo (por exemplo, `<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="51a3a-109">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="51a3a-110">Contido em</span><span class="sxs-lookup"><span data-stu-id="51a3a-110">Contained in</span></span>

[<span data-ttu-id="51a3a-111">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="51a3a-111">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="51a3a-112">Pode conter</span><span class="sxs-lookup"><span data-stu-id="51a3a-112">Can contain</span></span>

[<span data-ttu-id="51a3a-113">AppDomain</span><span class="sxs-lookup"><span data-stu-id="51a3a-113">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="51a3a-114">Comentários</span><span class="sxs-lookup"><span data-stu-id="51a3a-114">Remarks</span></span>

<span data-ttu-id="51a3a-115">Por padrão, o seu suplemento pode carregar qualquer página que esteja no mesmo domínio que o local especificado no elemento [SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="51a3a-115">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="51a3a-116">Para carregar páginas que não estejam no mesmo domínio do que o suplemento, especifique os domínios usando os elementos **AppDomains** e **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="51a3a-116">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="51a3a-117">Esse elemento não pode estar vazio.</span><span class="sxs-lookup"><span data-stu-id="51a3a-117">This element can't be empty.</span></span>
