---
title: Elemento AppDomains no arquivo de manifesto
description: ''
ms.date: 12/13/2018
ms.openlocfilehash: cc2f5ade0bdda214c85490f8e474b42f921edbe8
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433674"
---
# <a name="appdomains-element"></a><span data-ttu-id="e4e16-102">Elemento AppDomains</span><span class="sxs-lookup"><span data-stu-id="e4e16-102">AppDomains element</span></span>

<span data-ttu-id="e4e16-p101">Lista qualquer domínio além do domínio especificado no elemento SourceLocation que seu Suplemento do Office utilizará para carregar páginas. Para cada domínio adicional, especifique um elemento AppDomain.</span><span class="sxs-lookup"><span data-stu-id="e4e16-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="e4e16-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="e4e16-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e4e16-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="e4e16-106">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="e4e16-107">O valor de cada elemento **AppDomain** deve incluir o protocolo (por exemplo, `<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="e4e16-107">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="e4e16-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="e4e16-108">Contained in</span></span>

[<span data-ttu-id="e4e16-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="e4e16-109">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="e4e16-110">Pode conter</span><span class="sxs-lookup"><span data-stu-id="e4e16-110">Can contain</span></span>

[<span data-ttu-id="e4e16-111">AppDomain</span><span class="sxs-lookup"><span data-stu-id="e4e16-111">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="e4e16-112">Comentários</span><span class="sxs-lookup"><span data-stu-id="e4e16-112">Remarks</span></span>

<span data-ttu-id="e4e16-113">Por padrão, o seu suplemento pode carregar qualquer página que esteja no mesmo domínio que o local especificado no elemento [SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="e4e16-113">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="e4e16-114">Para carregar páginas que não estejam no mesmo domínio do que o suplemento, especifique os domínios usando os elementos **AppDomains** e **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="e4e16-114">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="e4e16-115">Esse elemento não pode estar vazio.</span><span class="sxs-lookup"><span data-stu-id="e4e16-115">This element can't be empty.</span></span>
