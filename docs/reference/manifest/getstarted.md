---
title: Elemento GetStarted no arquivo de manifesto
description: Fornece informações usadas pelo texto explicante que aparece quando o complemento é instalado no Word, Excel, PowerPoint e OneNote.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a637f3f9031d9f8e09d14f17f2095ca0647c4d50
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348682"
---
# <a name="getstarted-element"></a><span data-ttu-id="e450f-103">Elemento GetStarted</span><span class="sxs-lookup"><span data-stu-id="e450f-103">GetStarted element</span></span>

<span data-ttu-id="e450f-104">Fornece informações usadas pelo texto explicante que aparece quando o complemento é instalado no Word, Excel, PowerPoint e OneNote.</span><span class="sxs-lookup"><span data-stu-id="e450f-104">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote.</span></span> <span data-ttu-id="e450f-105">O elemento **GetStarted** é um elemento filho de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="e450f-105">The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="e450f-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="e450f-106">Child elements</span></span>

| <span data-ttu-id="e450f-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="e450f-107">Element</span></span>                       | <span data-ttu-id="e450f-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="e450f-108">Required</span></span> | <span data-ttu-id="e450f-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="e450f-109">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="e450f-110">Title</span><span class="sxs-lookup"><span data-stu-id="e450f-110">Title</span></span>](#title)               | <span data-ttu-id="e450f-111">Sim</span><span class="sxs-lookup"><span data-stu-id="e450f-111">Yes</span></span>      | <span data-ttu-id="e450f-112">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="e450f-112">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="e450f-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="e450f-113">Description</span></span>](#description)   | <span data-ttu-id="e450f-114">Sim</span><span class="sxs-lookup"><span data-stu-id="e450f-114">Yes</span></span>      | <span data-ttu-id="e450f-115">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="e450f-115">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="e450f-116">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="e450f-116">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="e450f-117">Sim</span><span class="sxs-lookup"><span data-stu-id="e450f-117">Yes</span></span>       | <span data-ttu-id="e450f-118">Uma URL para uma página que explica o suplemento em detalhes.</span><span class="sxs-lookup"><span data-stu-id="e450f-118">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="e450f-119">Título</span><span class="sxs-lookup"><span data-stu-id="e450f-119">Title</span></span> 

<span data-ttu-id="e450f-120">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="e450f-120">Required.</span></span> <span data-ttu-id="e450f-121">O título usado para o início do texto explicativo.</span><span class="sxs-lookup"><span data-stu-id="e450f-121">The title used for the top of the callout.</span></span> <span data-ttu-id="e450f-122">O **atributo resid** faz referência a uma ID válida no elemento **ShortStrings** na seção [Recursos](resources.md) e não pode ter mais de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e450f-122">The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="description"></a><span data-ttu-id="e450f-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="e450f-123">Description</span></span>

<span data-ttu-id="e450f-124">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="e450f-124">Required.</span></span> <span data-ttu-id="e450f-125">A descrição / conteúdo do corpo para o texto explicativo.</span><span class="sxs-lookup"><span data-stu-id="e450f-125">The description / body content for the callout.</span></span> <span data-ttu-id="e450f-126">O **atributo resid** faz referência a uma ID válida no elemento **LongStrings** na seção [Recursos](resources.md) e não pode ter mais de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e450f-126">The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="e450f-127">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="e450f-127">LearnMoreUrl</span></span>

<span data-ttu-id="e450f-128">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="e450f-128">Required.</span></span> <span data-ttu-id="e450f-129">A URL para uma página onde o usuário pode saber mais sobre o suplemento.</span><span class="sxs-lookup"><span data-stu-id="e450f-129">The URL to a page where the user can learn more about your add-in.</span></span> <span data-ttu-id="e450f-130">O **atributo resid** faz referência a uma ID válida no elemento **Urls** na seção [Recursos](resources.md) e não pode ter mais de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e450f-130">The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

> [!NOTE]
> <span data-ttu-id="e450f-131">**LearnMoreUrl** atualmente não é processado em clientes do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="e450f-131">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="e450f-132">Recomendamos que você adicione essa URL a todos os clientes para que a URL seja processada quando ficar disponível.</span><span class="sxs-lookup"><span data-stu-id="e450f-132">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="e450f-133">Confira também</span><span class="sxs-lookup"><span data-stu-id="e450f-133">See also</span></span>

<span data-ttu-id="e450f-134">Os exemplos de código a seguir usam o **elemento GetStarted.**</span><span class="sxs-lookup"><span data-stu-id="e450f-134">The following code samples use the **GetStarted** element.</span></span>

* [<span data-ttu-id="e450f-135">Suplemento Web do Excel para manipular formatação de tabelas e gráficos</span><span class="sxs-lookup"><span data-stu-id="e450f-135">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="e450f-136">JavaScript SpecKit para um Suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="e450f-136">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="e450f-137">Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e450f-137">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
