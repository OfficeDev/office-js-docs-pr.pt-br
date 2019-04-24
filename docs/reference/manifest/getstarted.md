---
title: Elemento GetStarted no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d9ebcba7881b388544eeb3e2c3028bff9bdcf9a6
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452078"
---
# <a name="getstarted-element"></a><span data-ttu-id="eefeb-102">Elemento GetStarted</span><span class="sxs-lookup"><span data-stu-id="eefeb-102">GetStarted element</span></span>

<span data-ttu-id="eefeb-p101">Fornece informações usadas pelo balão que aparece quando o suplemento está instalado em hosts do Word, do Excel, do PowerPoint e do OneNote. O elemento **GetStarted** é um elemento filho de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="eefeb-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="eefeb-105">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="eefeb-105">Child elements</span></span>

| <span data-ttu-id="eefeb-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="eefeb-106">Element</span></span>                       | <span data-ttu-id="eefeb-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="eefeb-107">Required</span></span> | <span data-ttu-id="eefeb-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="eefeb-108">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="eefeb-109">Title</span><span class="sxs-lookup"><span data-stu-id="eefeb-109">Title</span></span>](#title)               | <span data-ttu-id="eefeb-110">Sim</span><span class="sxs-lookup"><span data-stu-id="eefeb-110">Yes</span></span>      | <span data-ttu-id="eefeb-111">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="eefeb-111">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="eefeb-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="eefeb-112">Description</span></span>](#description)   | <span data-ttu-id="eefeb-113">Sim</span><span class="sxs-lookup"><span data-stu-id="eefeb-113">Yes</span></span>      | <span data-ttu-id="eefeb-114">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="eefeb-114">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="eefeb-115">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="eefeb-115">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="eefeb-116">Não</span><span class="sxs-lookup"><span data-stu-id="eefeb-116">No</span></span>       | <span data-ttu-id="eefeb-117">Uma URL para uma página que explica o suplemento em detalhes.</span><span class="sxs-lookup"><span data-stu-id="eefeb-117">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="eefeb-118">Título</span><span class="sxs-lookup"><span data-stu-id="eefeb-118">Title</span></span> 

<span data-ttu-id="eefeb-p102">Obrigatório. O título usado para o início do texto explicativo. O atributo **resid** faz referência a uma identificação válida no elemento **ShortStrings** na seção [Recursos](resources.md).</span><span class="sxs-lookup"><span data-stu-id="eefeb-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="eefeb-122">Descrição</span><span class="sxs-lookup"><span data-stu-id="eefeb-122">Description</span></span>

<span data-ttu-id="eefeb-p103">Obrigatório. A descrição / conteúdo do corpo para o texto explicativo. O atributo **resid** faz referência a uma identificação válida no elemento **LongStrings** na seção [Recursos](resources.md).</span><span class="sxs-lookup"><span data-stu-id="eefeb-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="eefeb-126">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="eefeb-126">LearnMoreUrl</span></span>

<span data-ttu-id="eefeb-p104">Obrigatório. A URL para uma página onde o usuário pode saber mais sobre o suplemento. O atributo **resid** faz referência a uma identificação válida no elemento **Urls** na seção [Recursos](resources.md).</span><span class="sxs-lookup"><span data-stu-id="eefeb-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="eefeb-130">**LearnMoreUrl** atualmente não é processado em clientes do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="eefeb-130">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="eefeb-131">Recomendamos que você adicione essa URL a todos os clientes para que a URL seja processada quando ficar disponível.</span><span class="sxs-lookup"><span data-stu-id="eefeb-131">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="eefeb-132">Confira também</span><span class="sxs-lookup"><span data-stu-id="eefeb-132">See also</span></span>

<span data-ttu-id="eefeb-133">Os exemplos de código a seguir utilizam o elemento **GetStarted**:</span><span class="sxs-lookup"><span data-stu-id="eefeb-133">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="eefeb-134">Suplemento Web do Excel para manipular formatação de tabelas e gráficos</span><span class="sxs-lookup"><span data-stu-id="eefeb-134">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="eefeb-135">JavaScript SpecKit para um Suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="eefeb-135">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="eefeb-136">Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="eefeb-136">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
