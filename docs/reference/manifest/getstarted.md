---
title: Elemento GetStarted no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 82fa1b9b62674adfb05c07536a7fdf2bbabf8f45
ms.sourcegitcommit: e5a5ec4ba32bacd0ccd13291b4e7f4bfc42901a3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/09/2019
ms.locfileid: "37429732"
---
# <a name="getstarted-element"></a><span data-ttu-id="15190-102">Elemento GetStarted</span><span class="sxs-lookup"><span data-stu-id="15190-102">GetStarted element</span></span>

<span data-ttu-id="15190-p101">Fornece informações usadas pelo balão que aparece quando o suplemento está instalado em hosts do Word, do Excel, do PowerPoint e do OneNote. O elemento **GetStarted** é um elemento filho de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="15190-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="15190-105">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="15190-105">Child elements</span></span>

| <span data-ttu-id="15190-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="15190-106">Element</span></span>                       | <span data-ttu-id="15190-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="15190-107">Required</span></span> | <span data-ttu-id="15190-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="15190-108">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="15190-109">Title</span><span class="sxs-lookup"><span data-stu-id="15190-109">Title</span></span>](#title)               | <span data-ttu-id="15190-110">Sim</span><span class="sxs-lookup"><span data-stu-id="15190-110">Yes</span></span>      | <span data-ttu-id="15190-111">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="15190-111">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="15190-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="15190-112">Description</span></span>](#description)   | <span data-ttu-id="15190-113">Sim</span><span class="sxs-lookup"><span data-stu-id="15190-113">Yes</span></span>      | <span data-ttu-id="15190-114">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="15190-114">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="15190-115">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="15190-115">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="15190-116">Sim</span><span class="sxs-lookup"><span data-stu-id="15190-116">Yes</span></span>       | <span data-ttu-id="15190-117">Uma URL para uma página que explica o suplemento em detalhes.</span><span class="sxs-lookup"><span data-stu-id="15190-117">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="15190-118">Cargo</span><span class="sxs-lookup"><span data-stu-id="15190-118">Title</span></span> 

<span data-ttu-id="15190-p102">Obrigatório. O título usado para o início do texto explicativo. O atributo **resid** faz referência a uma identificação válida no elemento **ShortStrings** na seção [Recursos](resources.md).</span><span class="sxs-lookup"><span data-stu-id="15190-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="15190-122">Descrição</span><span class="sxs-lookup"><span data-stu-id="15190-122">Description</span></span>

<span data-ttu-id="15190-p103">Obrigatório. A descrição / conteúdo do corpo para o texto explicativo. O atributo **resid** faz referência a uma identificação válida no elemento **LongStrings** na seção [Recursos](resources.md).</span><span class="sxs-lookup"><span data-stu-id="15190-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="15190-126">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="15190-126">LearnMoreUrl</span></span>

<span data-ttu-id="15190-p104">Obrigatório. A URL para uma página onde o usuário pode saber mais sobre o suplemento. O atributo **resid** faz referência a uma identificação válida no elemento **Urls** na seção [Recursos](resources.md).</span><span class="sxs-lookup"><span data-stu-id="15190-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="15190-130">**LearnMoreUrl** atualmente não é processado em clientes do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="15190-130">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="15190-131">Recomendamos que você adicione essa URL a todos os clientes para que a URL seja processada quando ficar disponível.</span><span class="sxs-lookup"><span data-stu-id="15190-131">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="15190-132">Confira também</span><span class="sxs-lookup"><span data-stu-id="15190-132">See also</span></span>

<span data-ttu-id="15190-133">Os exemplos de código a seguir utilizam o elemento **GetStarted**:</span><span class="sxs-lookup"><span data-stu-id="15190-133">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="15190-134">Suplemento Web do Excel para manipular formatação de tabelas e gráficos</span><span class="sxs-lookup"><span data-stu-id="15190-134">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="15190-135">JavaScript SpecKit para um Suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="15190-135">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="15190-136">Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="15190-136">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
