---
title: Office UI Fabric em suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 04964d5864eea4a960f7b57e5df6f7bd7c844fde
ms.sourcegitcommit: 4e4f7c095e8f33b06bd8a02534ee901125eb1d17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/28/2018
ms.locfileid: "20084067"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="d16dc-102">Office UI Fabric em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="d16dc-102">Office UI Fabric in Office Add-ins</span></span> 

<span data-ttu-id="d16dc-p101">O Office UI Fabric é uma estrutura de front-end JavaScript destinada à criação de experiências de usuário para Office e Office 365. O Fabric fornece componentes com foco em efeitos visuais que você pode estender, reformular e usar no suplemento do Office. Como o Fabric usa a linguagem de design da Microsoft, os componentes da experiência de usuário do Fabric são semelhantes a uma extensão natural do Office.</span><span class="sxs-lookup"><span data-stu-id="d16dc-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span> 

<span data-ttu-id="d16dc-p102">Se estiver criando um suplemento, recomendamos usar o Office UI Fabric para criar a experiência de usuário. O uso do Office UI Fabric é opcional.</span><span class="sxs-lookup"><span data-stu-id="d16dc-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="d16dc-108">As seções a seguir explicam como começar a usar o Fabric para atender às suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="d16dc-108">The following sections explain how to get started using Fabric to meet your requirements.</span></span> 

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="d16dc-109">Uso do Fabric Core: ícones, fontes, cores</span><span class="sxs-lookup"><span data-stu-id="d16dc-109">Use Fabric Core: icons, fonts, colors</span></span>
<span data-ttu-id="d16dc-p103">O Fabric Core contém os elementos principais da linguagem de design, como ícones, cores, tipo e grade. O Fabric Core é independente de estrutura. Tanto o Fabric JS como o Fabric React usam o Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="d16dc-p103">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid. Fabric core is framework independent. Both Fabric React and Fabric JS use Fabric Core.</span></span>

<span data-ttu-id="d16dc-113">Para começar a usar o Fabric Core:</span><span class="sxs-lookup"><span data-stu-id="d16dc-113">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="d16dc-114">Adicione a referência da CDN ao HTML da sua página.</span><span class="sxs-lookup"><span data-stu-id="d16dc-114">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. <span data-ttu-id="d16dc-115">Use ícones e fontes do Fabric.</span><span class="sxs-lookup"><span data-stu-id="d16dc-115">Use Fabric icons and fonts.</span></span> 

    <span data-ttu-id="d16dc-p104">Para usar um ícone do Fabric, inclua o elemento "i" na sua página e, em seguida, faça referência às classes apropriadas. Para controlar o tamanho do ícone, você pode alterar o tamanho da fonte. Por exemplo, o código a seguir mostra como criar um ícone de tabela muito grande que usa a cor themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="d16dc-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span> 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="d16dc-p105">Para localizar mais ícones disponíveis no Office UI Fabric, use o recurso de pesquisa na página [Ícones](https://dev.office.com/fabric#/styles/icons). Quando encontrar um ícone para usar no suplemento, não deixe de adicionar um prefixo ao nome do ícone com `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="d16dc-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://dev.office.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span> 

    <span data-ttu-id="d16dc-121">Para saber mais sobre os tamanhos de fonte e as cores disponíveis no Office UI Fabric, confira [Tipografia](https://dev.office.com/fabric#/styles/typography) e [Cores](https://dev.office.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="d16dc-121">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://dev.office.com/fabric#/styles/typography) and [Colors](https://dev.office.com/fabric#/styles/colors).</span></span>
 
## <a name="use-fabric-components"></a><span data-ttu-id="d16dc-122">Uso dos componentes do Fabric</span><span class="sxs-lookup"><span data-stu-id="d16dc-122">Use Fabric Components</span></span> 
<span data-ttu-id="d16dc-123">O Fabric oferece uma variedade de componentes da experiência do usuário que você pode usar para criar o suplemento. Alguns desses componentes incluem:</span><span class="sxs-lookup"><span data-stu-id="d16dc-123">Fabric provides a variety of UX components that you can use to build your add-in, including the following types of components:</span></span>

- <span data-ttu-id="d16dc-124">Componentes de entrada – por exemplo, botão, caixa de seleção e alternância</span><span class="sxs-lookup"><span data-stu-id="d16dc-124">Input components - for example, Button, Checkbox, and Toggle</span></span>
- <span data-ttu-id="d16dc-125">Componentes de navegação – por exemplo, tabela dinâmica e navegação de trilha</span><span class="sxs-lookup"><span data-stu-id="d16dc-125">Navigation components - for example, Pivot Breadcrumb</span></span>
- <span data-ttu-id="d16dc-126">Componentes de notificação – por exemplo, MessageBar e balão</span><span class="sxs-lookup"><span data-stu-id="d16dc-126">Notification components - for example, MessageBar and Callout</span></span>  

<span data-ttu-id="d16dc-127">Nem todos os componentes do Fabric são recomendados para uso em suplementos. Aqui está uma lista de componentes do Fabric React UX que recomendamos para suplementos:</span><span class="sxs-lookup"><span data-stu-id="d16dc-127">Not all Fabric components are recommended for use in add-ins. Here is a list of Fabric React UX components that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="d16dc-128">Navegação de trilha</span><span class="sxs-lookup"><span data-stu-id="d16dc-128">Breadcrumb</span></span>](https://developer.microsoft.com/en-us/fabric#/components/breadcrumb)
- [<span data-ttu-id="d16dc-129">Botão</span><span class="sxs-lookup"><span data-stu-id="d16dc-129">Button</span></span>](https://developer.microsoft.com/en-us/fabric#/components/button)
- [<span data-ttu-id="d16dc-130">Caixa de seleção</span><span class="sxs-lookup"><span data-stu-id="d16dc-130">Checkbox</span></span>](https://developer.microsoft.com/en-us/fabric#/components/checkbox)
- [<span data-ttu-id="d16dc-131">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="d16dc-131">ChoiceGroup</span></span>](https://developer.microsoft.com/en-us/fabric#/components/choicegroup)
- [<span data-ttu-id="d16dc-132">Lista suspensa</span><span class="sxs-lookup"><span data-stu-id="d16dc-132">Dropdown</span></span>](https://developer.microsoft.com/en-us/fabric#/components/dropdown)
- [<span data-ttu-id="d16dc-133">Rótulo</span><span class="sxs-lookup"><span data-stu-id="d16dc-133">Label</span></span>](https://developer.microsoft.com/en-us/fabric#/components/label)
- [<span data-ttu-id="d16dc-134">Lista</span><span class="sxs-lookup"><span data-stu-id="d16dc-134">List</span></span>](https://developer.microsoft.com/en-us/fabric#/components/list)
- [<span data-ttu-id="d16dc-135">Tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="d16dc-135">Pivot</span></span>](https://developer.microsoft.com/en-us/fabric#/components/pivot)
- [<span data-ttu-id="d16dc-136">TextField</span><span class="sxs-lookup"><span data-stu-id="d16dc-136">TextField</span></span>](https://developer.microsoft.com/en-us/fabric#/components/textfield)
- [<span data-ttu-id="d16dc-137">Alternância</span><span class="sxs-lookup"><span data-stu-id="d16dc-137">Toggle</span></span>](https://developer.microsoft.com/en-us/fabric#/components/toggle)

<span data-ttu-id="d16dc-p106">Você pode usar diferentes estruturas do JavaScript, como Angular ou React, para criar o suplemento. Para começar a usar componentes do Fabric com sua estrutura, confira os recursos a seguir.</span><span class="sxs-lookup"><span data-stu-id="d16dc-p106">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="d16dc-140">**Estrutura**</span><span class="sxs-lookup"><span data-stu-id="d16dc-140">**Framework**</span></span>|<span data-ttu-id="d16dc-141">**Exemplo**</span><span class="sxs-lookup"><span data-stu-id="d16dc-141">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="d16dc-142">**React**</span><span class="sxs-lookup"><span data-stu-id="d16dc-142">**React**</span></span>|[<span data-ttu-id="d16dc-143">Uso do Office UI Fabric React em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="d16dc-143">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="d16dc-144">**Angular**</span><span class="sxs-lookup"><span data-stu-id="d16dc-144">**Angular**</span></span>| <span data-ttu-id="d16dc-145">Confira [ngOfficeUIFabric](http://ngofficeuifabric.com/), que é um projeto comunitário com diretivas do Angular 1.5, e [Considere a possibilidade de dispor componentes do Fabric com componentes do Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span><span class="sxs-lookup"><span data-stu-id="d16dc-145">See [ngOfficeUIFabric](http://ngofficeuifabric.com/), which is a community project with Angular 1.5 directives, and [Consider wrapping Fabric components with Angular 2 components](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span></span>|
