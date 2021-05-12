---
title: Fabric Core em Office de complementos
description: Obter uma visão geral de como usar o Fabric Core e os componentes da interface do usuário do Fabric em Office de complementos.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: e93efaea55841cc3bb6fa79ea1d1bbcaa76a4d05
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330196"
---
# <a name="fabric-core-in-office-add-ins"></a><span data-ttu-id="6f5d8-103">Fabric Core em Office de complementos</span><span class="sxs-lookup"><span data-stu-id="6f5d8-103">Fabric Core in Office Add-ins</span></span>

<span data-ttu-id="6f5d8-104">Fabric Core é uma coleção open-source de classes CSS e mixins SASS que se destinam a ser usadas em React *Office* Add-ins. O Fabric Core contém elementos básicos da linguagem de design da interface do usuário fluente, como ícones, cores, tipos e grades.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-104">Fabric Core is an open-source collection of CSS classes and SASS mixins that's *intended for use in non-React* Office Add-ins. Fabric Core contains basic elements of the Fluent UI design language such as icons, colors, typefaces, and grids.</span></span> <span data-ttu-id="6f5d8-105">O Fabric Core é independente da estrutura, portanto, pode ser usado com qualquer aplicativo de página única ou qualquer estrutura de interface do usuário web do lado do servidor.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-105">Fabric Core is framework independent, so it can be used with any single-page application or any server-side web UI framework.</span></span> <span data-ttu-id="6f5d8-106">(Chama-se "Fabric Core" em vez de "Fluent Core" por motivos históricos.)</span><span class="sxs-lookup"><span data-stu-id="6f5d8-106">(It's called "Fabric Core" instead of "Fluent Core" for historical reasons.)</span></span>

<span data-ttu-id="6f5d8-107">Se a interface do usuário do seu React não for baseada em React, você também poderá usar um conjunto de componentes que não React.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-107">If your add-in's UI is not React-based, you can also make use of a set of non-React components.</span></span> <span data-ttu-id="6f5d8-108">Consulte [Usar Office UI Fabric componentes JS](#use-office-ui-fabric-js-components).</span><span class="sxs-lookup"><span data-stu-id="6f5d8-108">See [Use Office UI Fabric JS components](#use-office-ui-fabric-js-components).</span></span>

> [!NOTE]
> <span data-ttu-id="6f5d8-109">Este artigo descreve o uso do Fabric Core no contexto de Office de complementos. Mas também é usado em uma ampla variedade de Microsoft 365 aplicativos e extensões.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-109">This article describes the use of Fabric Core in the context of Office Add-ins. But it's also used in a wide range of Microsoft 365 apps and extensions.</span></span> <span data-ttu-id="6f5d8-110">Para obter mais informações, [consulte Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) and the open source repo Office UI Fabric [Core](https://github.com/OfficeDev/office-ui-fabric-core).</span><span class="sxs-lookup"><span data-stu-id="6f5d8-110">For more information, see [Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) and the open source repo [Office UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core).</span></span>

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="6f5d8-111">Uso do Fabric Core: ícones, fontes, cores</span><span class="sxs-lookup"><span data-stu-id="6f5d8-111">Use Fabric Core: icons, fonts, colors</span></span>

<span data-ttu-id="6f5d8-112">Para começar a usar o Fabric Core:</span><span class="sxs-lookup"><span data-stu-id="6f5d8-112">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="6f5d8-113">Adicione a referência da CDN ao HTML da sua página.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-113">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. <span data-ttu-id="6f5d8-114">Use ícones e fontes do Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-114">Use Fabric Core icons and fonts.</span></span>

    <span data-ttu-id="6f5d8-115">Para usar um ícone do Fabric Core, inclua o elemento "i" em sua página e, em seguida, fazer referência às classes apropriadas.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-115">To use a Fabric Core icon, include the "i" element on your page, and then reference the appropriate classes.</span></span> <span data-ttu-id="6f5d8-116">Para controlar o tamanho do ícone, você pode alterar o tamanho da fonte.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-116">You can control the size of the icon by changing the font size.</span></span> <span data-ttu-id="6f5d8-117">Por exemplo, o código a seguir mostra como criar um ícone de tabela muito grande que usa a cor themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="6f5d8-117">For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span>

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="6f5d8-118">Para obter instruções mais detalhadas, consulte [Fluent UI Icons](https://developer.microsoft.com/fluentui#/styles/web/icons).</span><span class="sxs-lookup"><span data-stu-id="6f5d8-118">For more detailed instructions, see [Fluent UI Icons](https://developer.microsoft.com/fluentui#/styles/web/icons).</span></span> <span data-ttu-id="6f5d8-119">Para encontrar mais ícones disponíveis no Fabric Core, use o recurso de pesquisa nessa página.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-119">To find more icons that are available in Fabric Core, use the search feature on that page.</span></span> <span data-ttu-id="6f5d8-120">Quando encontrar um ícone para usar no suplemento, não deixe de adicionar um prefixo ao nome do ícone com `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-120">When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span>

    <span data-ttu-id="6f5d8-121">Para obter informações sobre tamanhos de fonte e cores disponíveis no Fabric Core, consulte [Typography](https://developer.microsoft.com/fluentui#/styles/web/typography) and the **Colors** table of contents at [Colors](https://developer.microsoft.com/fluentui#/styles/web/colors).</span><span class="sxs-lookup"><span data-stu-id="6f5d8-121">For information about font sizes and colors that are available in Fabric Core, see [Typography](https://developer.microsoft.com/fluentui#/styles/web/typography) and the **Colors** table of contents at [Colors](https://developer.microsoft.com/fluentui#/styles/web/colors).</span></span>

<span data-ttu-id="6f5d8-122">Exemplos são incluídos nos [Exemplos](#samples) posteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-122">Examples are included in the [Samples](#samples) later in this article.</span></span>

## <a name="use-office-ui-fabric-js-components"></a><span data-ttu-id="6f5d8-123">Usar Office UI Fabric JS</span><span class="sxs-lookup"><span data-stu-id="6f5d8-123">Use Office UI Fabric JS components</span></span>

<span data-ttu-id="6f5d8-124">Os complementos com UIs não React também podem usar qualquer um dos muitos componentes do [Office UI Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js), incluindo botões, caixas de diálogo, seladores e muito mais.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-124">Add-ins with non-React UIs can also use any of the many components from [Office UI Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js), including buttons, dialogs, pickers, and more.</span></span> <span data-ttu-id="6f5d8-125">Consulte o readme do repo para obter instruções.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-125">See the readme of the repo for instructions.</span></span>

<span data-ttu-id="6f5d8-126">Exemplos são incluídos nos [Exemplos](#samples) posteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-126">Examples are included in the [Samples](#samples) later in this article.</span></span>

## <a name="samples"></a><span data-ttu-id="6f5d8-127">Exemplos</span><span class="sxs-lookup"><span data-stu-id="6f5d8-127">Samples</span></span>

<span data-ttu-id="6f5d8-128">Os seguintes exemplos de complementos usam o Fabric Core e/ou Office UI Fabric componentes JS.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-128">The following sample add-ins use Fabric Core and/or Office UI Fabric JS components.</span></span> <span data-ttu-id="6f5d8-129">Algumas dessas repos são arquivadas, o que significa que elas não estão mais sendo atualizadas com correções de bugs ou de segurança, mas você ainda pode usá-las para aprender a usar componentes do Fabric Core e da interface do usuário do Fabric.</span><span class="sxs-lookup"><span data-stu-id="6f5d8-129">Some of these repos are archived, meaning that they are no longer being updated with bug or security fixes, but you can still use them to learn how to use Fabric Core and Fabric UI components.</span></span>

- [<span data-ttu-id="6f5d8-130">Excel Add-in JavaScript SalesTracker</span><span class="sxs-lookup"><span data-stu-id="6f5d8-130">Excel Add-in JavaScript SalesTracker</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
- [<span data-ttu-id="6f5d8-131">Excel SalesLeads de complemento</span><span class="sxs-lookup"><span data-stu-id="6f5d8-131">Excel Add-in SalesLeads</span></span>](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)
- [<span data-ttu-id="6f5d8-132">Excel Tendências de despesas de woodgrove do add-in</span><span class="sxs-lookup"><span data-stu-id="6f5d8-132">Excel Add-in WoodGrove Expense Trends</span></span>](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [<span data-ttu-id="6f5d8-133">Excel Content Add-in Humongous Insurance</span><span class="sxs-lookup"><span data-stu-id="6f5d8-133">Excel Content Add-in Humongous Insurance</span></span>](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
- [<span data-ttu-id="6f5d8-134">Office Exemplo de interface do usuário do Fabric do add-in</span><span class="sxs-lookup"><span data-stu-id="6f5d8-134">Office Add-in Fabric UI Sample</span></span>](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [<span data-ttu-id="6f5d8-135">Office-Add-in-UX-Design-Patterns-Code</span><span class="sxs-lookup"><span data-stu-id="6f5d8-135">Office-Add-in-UX-Design-Patterns-Code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="6f5d8-136">Outlook Add-in GifMe</span><span class="sxs-lookup"><span data-stu-id="6f5d8-136">Outlook Add-in GifMe</span></span>](https://github.com/OfficeDev/Outlook-Add-in-GifMe)
- [<span data-ttu-id="6f5d8-137">PowerPoint Add-in Microsoft Graph ASPNET InsertChart</span><span class="sxs-lookup"><span data-stu-id="6f5d8-137">PowerPoint Add-in Microsoft Graph ASPNET InsertChart</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [<span data-ttu-id="6f5d8-138">Word Add-in Angular2 StyleChecker</span><span class="sxs-lookup"><span data-stu-id="6f5d8-138">Word Add-in Angular2 StyleChecker</span></span>](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)
- [<span data-ttu-id="6f5d8-139">Word Add-in JS Redact</span><span class="sxs-lookup"><span data-stu-id="6f5d8-139">Word Add-in JS Redact</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [<span data-ttu-id="6f5d8-140">Word Add-in MarkdownConversion</span><span class="sxs-lookup"><span data-stu-id="6f5d8-140">Word Add-in MarkdownConversion</span></span>](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion)
