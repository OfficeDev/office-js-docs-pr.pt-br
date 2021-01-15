---
title: Elemento GetStarted no arquivo de manifesto
description: Fornece informações usadas pelo texto explicativo que aparece quando o suplemento é instalado no Word, Excel, PowerPoint e OneNote.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0ad6196dc45e4ea06c2b43ac5da66a560ab0b899
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771411"
---
# <a name="getstarted-element"></a><span data-ttu-id="5a18a-103">Elemento GetStarted</span><span class="sxs-lookup"><span data-stu-id="5a18a-103">GetStarted element</span></span>

<span data-ttu-id="5a18a-104">Fornece informações usadas pelo texto explicativo que aparece quando o suplemento é instalado no Word, Excel, PowerPoint e OneNote.</span><span class="sxs-lookup"><span data-stu-id="5a18a-104">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote.</span></span> <span data-ttu-id="5a18a-105">O elemento **GetStarted** é um elemento filho de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="5a18a-105">The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="5a18a-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="5a18a-106">Child elements</span></span>

| <span data-ttu-id="5a18a-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="5a18a-107">Element</span></span>                       | <span data-ttu-id="5a18a-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="5a18a-108">Required</span></span> | <span data-ttu-id="5a18a-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="5a18a-109">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="5a18a-110">Title</span><span class="sxs-lookup"><span data-stu-id="5a18a-110">Title</span></span>](#title)               | <span data-ttu-id="5a18a-111">Sim</span><span class="sxs-lookup"><span data-stu-id="5a18a-111">Yes</span></span>      | <span data-ttu-id="5a18a-112">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="5a18a-112">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="5a18a-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="5a18a-113">Description</span></span>](#description)   | <span data-ttu-id="5a18a-114">Sim</span><span class="sxs-lookup"><span data-stu-id="5a18a-114">Yes</span></span>      | <span data-ttu-id="5a18a-115">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="5a18a-115">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="5a18a-116">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="5a18a-116">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="5a18a-117">Sim</span><span class="sxs-lookup"><span data-stu-id="5a18a-117">Yes</span></span>       | <span data-ttu-id="5a18a-118">Uma URL para uma página que explica o suplemento em detalhes.</span><span class="sxs-lookup"><span data-stu-id="5a18a-118">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="5a18a-119">Título</span><span class="sxs-lookup"><span data-stu-id="5a18a-119">Title</span></span> 

<span data-ttu-id="5a18a-120">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="5a18a-120">Required.</span></span> <span data-ttu-id="5a18a-121">O título usado para o início do texto explicativo.</span><span class="sxs-lookup"><span data-stu-id="5a18a-121">The title used for the top of the callout.</span></span> <span data-ttu-id="5a18a-122">O atributo **Resid** faz referência a uma identificação válida no elemento **ShortStrings** na seção [recursos](resources.md) e não pode ter mais de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="5a18a-122">The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="description"></a><span data-ttu-id="5a18a-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="5a18a-123">Description</span></span>

<span data-ttu-id="5a18a-124">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="5a18a-124">Required.</span></span> <span data-ttu-id="5a18a-125">A descrição / conteúdo do corpo para o texto explicativo.</span><span class="sxs-lookup"><span data-stu-id="5a18a-125">The description / body content for the callout.</span></span> <span data-ttu-id="5a18a-126">O atributo **Resid** faz referência a uma identificação válida no elemento **LongStrings** na seção [recursos](resources.md) e não pode ter mais de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="5a18a-126">The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="5a18a-127">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="5a18a-127">LearnMoreUrl</span></span>

<span data-ttu-id="5a18a-128">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="5a18a-128">Required.</span></span> <span data-ttu-id="5a18a-129">A URL para uma página onde o usuário pode saber mais sobre o suplemento.</span><span class="sxs-lookup"><span data-stu-id="5a18a-129">The URL to a page where the user can learn more about your add-in.</span></span> <span data-ttu-id="5a18a-130">O atributo **Resid** faz referência a uma identificação válida no elemento **URLs** na seção [recursos](resources.md) e não pode ter mais de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="5a18a-130">The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section and can be no more than 32 characters.</span></span>

> [!NOTE]
> <span data-ttu-id="5a18a-131">**LearnMoreUrl** atualmente não é processado em clientes do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="5a18a-131">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="5a18a-132">Recomendamos que você adicione essa URL a todos os clientes para que a URL seja processada quando ficar disponível.</span><span class="sxs-lookup"><span data-stu-id="5a18a-132">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="5a18a-133">Confira também</span><span class="sxs-lookup"><span data-stu-id="5a18a-133">See also</span></span>

<span data-ttu-id="5a18a-134">Os exemplos de código a seguir utilizam o elemento **GetStarted**:</span><span class="sxs-lookup"><span data-stu-id="5a18a-134">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="5a18a-135">Suplemento Web do Excel para manipular formatação de tabelas e gráficos</span><span class="sxs-lookup"><span data-stu-id="5a18a-135">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="5a18a-136">JavaScript SpecKit para um Suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="5a18a-136">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="5a18a-137">Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5a18a-137">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
