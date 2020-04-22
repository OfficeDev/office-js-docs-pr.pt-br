---
title: Office UI Fabric em suplementos do Office 
description: Obtenha uma visão geral de como usar os componentes do Office UI Fabric em suplementos do Office.
ms.date: 04/20/2020
localization_priority: Normal
ms.openlocfilehash: 5c504de14ee97ff740a80dc7608ae636ff8080ca
ms.sourcegitcommit: 79c55e59294e220bd21a5006080f72acf3ec0a3f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/21/2020
ms.locfileid: "43581908"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="ee5e5-103">Office UI Fabric em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ee5e5-103">Office UI Fabric in Office Add-ins</span></span> 

<span data-ttu-id="ee5e5-p101">O Office UI Fabric é uma estrutura de front-end JavaScript destinada à criação de experiências de usuário para Office e Office 365. O Fabric fornece componentes com foco em efeitos visuais que você pode estender, reformular e usar no suplemento do Office. Como o Fabric usa a linguagem de design da Microsoft, os componentes da experiência de usuário do Fabric são semelhantes a uma extensão natural do Office.</span><span class="sxs-lookup"><span data-stu-id="ee5e5-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span> 

<span data-ttu-id="ee5e5-p102">Se estiver criando um suplemento, recomendamos usar o Office UI Fabric para criar a experiência de usuário. O uso do Office UI Fabric é opcional.</span><span class="sxs-lookup"><span data-stu-id="ee5e5-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="ee5e5-109">As seções a seguir explicam como começar a usar o Fabric para atender às suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="ee5e5-109">The following sections explain how to get started using Fabric to meet your requirements.</span></span> 

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="ee5e5-110">Uso do Fabric Core: ícones, fontes, cores</span><span class="sxs-lookup"><span data-stu-id="ee5e5-110">Use Fabric Core: icons, fonts, colors</span></span>
<span data-ttu-id="ee5e5-111">O Fabric Core contém os elementos principais da linguagem de design, como ícones, cores, tipo e grade.</span><span class="sxs-lookup"><span data-stu-id="ee5e5-111">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid.</span></span><span data-ttu-id="ee5e5-112"> O Fabric Core é independente de estrutura.</span><span class="sxs-lookup"><span data-stu-id="ee5e5-112"> Fabric core is framework independent.</span></span> <span data-ttu-id="ee5e5-113">O Fabric Core é usado pelo Fabric React e incluído nele.</span><span class="sxs-lookup"><span data-stu-id="ee5e5-113">Fabric Core is used by and included with Fabric React.</span></span>

<span data-ttu-id="ee5e5-114">Para começar a usar o Fabric Core:</span><span class="sxs-lookup"><span data-stu-id="ee5e5-114">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="ee5e5-115">Adicione a referência da CDN ao HTML da sua página.</span><span class="sxs-lookup"><span data-stu-id="ee5e5-115">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```   
    
2. <span data-ttu-id="ee5e5-116">Use ícones e fontes do Fabric.</span><span class="sxs-lookup"><span data-stu-id="ee5e5-116">Use Fabric icons and fonts.</span></span> 

    <span data-ttu-id="ee5e5-p104">Para usar um ícone do Fabric, inclua o elemento "i" na sua página e, em seguida, faça referência às classes apropriadas. Para controlar o tamanho do ícone, você pode alterar o tamanho da fonte. Por exemplo, o código a seguir mostra como criar um ícone de tabela muito grande que usa a cor themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="ee5e5-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span> 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="ee5e5-p105">Para localizar mais ícones disponíveis no Office UI Fabric, use o recurso de pesquisa na página [Ícones](https://developer.microsoft.com/fabric#/styles/icons). Quando encontrar um ícone para usar no suplemento, não deixe de adicionar um prefixo ao nome do ícone com `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="ee5e5-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://developer.microsoft.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span> 

    <span data-ttu-id="ee5e5-122">Para saber mais sobre os tamanhos de fonte e as cores disponíveis no Office UI Fabric, confira [Tipografia](https://developer.microsoft.com/fabric#/styles/typography) e [Cores](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="ee5e5-122">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://developer.microsoft.com/fabric#/styles/typography) and [Colors](https://developer.microsoft.com/fabric#/styles/colors).</span></span>
 
## <a name="use-fabric-components"></a><span data-ttu-id="ee5e5-123">Uso dos componentes do Fabric</span><span class="sxs-lookup"><span data-stu-id="ee5e5-123">Use Fabric Components</span></span> 
<span data-ttu-id="ee5e5-124">O Fabric oferece uma variedade de componentes da experiência do usuário que você pode usar para criar o suplemento. Alguns desses componentes incluem:</span><span class="sxs-lookup"><span data-stu-id="ee5e5-124">Fabric provides a variety of UX components that you can use to build your add-in, including the following types of components:</span></span>

- <span data-ttu-id="ee5e5-125">Componentes de entrada – por exemplo, botão, caixa de seleção e alternância</span><span class="sxs-lookup"><span data-stu-id="ee5e5-125">Input components - for example, Button, Checkbox, and Toggle</span></span>
- <span data-ttu-id="ee5e5-126">Componentes de navegação – por exemplo, dinâmico e trilha</span><span class="sxs-lookup"><span data-stu-id="ee5e5-126">Navigation components - for example, Pivot and Breadcrumb</span></span>
- <span data-ttu-id="ee5e5-127">Componentes de notificação – por exemplo, MessageBar e balão</span><span class="sxs-lookup"><span data-stu-id="ee5e5-127">Notification components - for example, MessageBar and Callout</span></span>  

<span data-ttu-id="ee5e5-128">Nem todos os componentes do Fabric são recomendados para usar em suplementos. Aqui está uma lista de componentes de experiência de usuário do Fabric React que recomendamos para uso em um suplemento:</span><span class="sxs-lookup"><span data-stu-id="ee5e5-128">Not all Fabric components are recommended for use in add-ins. Here is a list of Fabric React UX components that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="ee5e5-129">Navegação estrutural</span><span class="sxs-lookup"><span data-stu-id="ee5e5-129">Breadcrumb</span></span>](https://developer.microsoft.com/fabric#/components/breadcrumb)
- [<span data-ttu-id="ee5e5-130">Botão</span><span class="sxs-lookup"><span data-stu-id="ee5e5-130">Button</span></span>](https://developer.microsoft.com/fabric#/components/button)
- [<span data-ttu-id="ee5e5-131">Caixa de seleção</span><span class="sxs-lookup"><span data-stu-id="ee5e5-131">Checkbox</span></span>](https://developer.microsoft.com/fabric#/components/checkbox)
- [<span data-ttu-id="ee5e5-132">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="ee5e5-132">ChoiceGroup</span></span>](https://developer.microsoft.com/fabric#/components/choicegroup)
- [<span data-ttu-id="ee5e5-133">Lista suspensa</span><span class="sxs-lookup"><span data-stu-id="ee5e5-133">Dropdown</span></span>](https://developer.microsoft.com/fabric#/components/dropdown)
- [<span data-ttu-id="ee5e5-134">Rótulo</span><span class="sxs-lookup"><span data-stu-id="ee5e5-134">Label</span></span>](https://developer.microsoft.com/fabric#/components/label)
- [<span data-ttu-id="ee5e5-135">Lista</span><span class="sxs-lookup"><span data-stu-id="ee5e5-135">List</span></span>](https://developer.microsoft.com/fabric#/components/list)
- [<span data-ttu-id="ee5e5-136">Tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="ee5e5-136">Pivot</span></span>](https://developer.microsoft.com/fabric#/components/pivot)
- [<span data-ttu-id="ee5e5-137">Campo de texto</span><span class="sxs-lookup"><span data-stu-id="ee5e5-137">TextField</span></span>](https://developer.microsoft.com/fabric#/components/textfield)
- [<span data-ttu-id="ee5e5-138">Alternância</span><span class="sxs-lookup"><span data-stu-id="ee5e5-138">Toggle</span></span>](https://developer.microsoft.com/fabric#/components/toggle)

<span data-ttu-id="ee5e5-p106">Você pode usar diferentes estruturas do JavaScript, como Angular ou React, para criar o suplemento. Para começar a usar componentes do Fabric com sua estrutura, confira os recursos a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee5e5-p106">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="ee5e5-141">**Framework**</span><span class="sxs-lookup"><span data-stu-id="ee5e5-141">**Framework**</span></span>|<span data-ttu-id="ee5e5-142">**Exemplo**</span><span class="sxs-lookup"><span data-stu-id="ee5e5-142">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="ee5e5-143">**React**</span><span class="sxs-lookup"><span data-stu-id="ee5e5-143">**React**</span></span>|[<span data-ttu-id="ee5e5-144">Uso do Office UI Fabric React em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ee5e5-144">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="ee5e5-145">**Angular**</span><span class="sxs-lookup"><span data-stu-id="ee5e5-145">**Angular**</span></span>| [<span data-ttu-id="ee5e5-146">Considere quebrar componentes do Fabric com componentes angulares 2</span><span class="sxs-lookup"><span data-stu-id="ee5e5-146">Consider wrapping Fabric components with Angular 2 components</span></span>](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
