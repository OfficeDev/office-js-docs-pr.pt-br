---
title: Office UI Fabric em suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 8fafe8a68c477868c12bff61c7f9ff23fc7314e0
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="2b52f-102">Office UI Fabric em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="2b52f-102">Office UI Fabric in Office Add-ins</span></span> 

<span data-ttu-id="2b52f-p101">O Office UI Fabric ? uma estrutura de front-end JavaScript destinada ? cria??o de experi?ncias de usu?rio para Office e Office 365. O Fabric fornece componentes com foco em efeitos visuais que voc? pode estender, reformular e usar no suplemento do Office. Como o Fabric usa a linguagem de design da Microsoft, os componentes da experi?ncia de usu?rio do Fabric s?o semelhantes a uma extens?o natural do Office.</span><span class="sxs-lookup"><span data-stu-id="2b52f-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span> 

<span data-ttu-id="2b52f-p102">Se estiver criando um suplemento, recomendamos usar o Office UI Fabric para criar a experi?ncia de usu?rio. O uso do Office UI Fabric ? opcional.</span><span class="sxs-lookup"><span data-stu-id="2b52f-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="2b52f-108">As se??es a seguir explicam como come?ar a usar o Fabric para atender ?s suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="2b52f-108">The following sections explain how to get started using Fabric to meet your requirements.</span></span> 

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="2b52f-109">Uso do Fabric Core: ?cones, fontes, cores</span><span class="sxs-lookup"><span data-stu-id="2b52f-109">Use Fabric Core: icons, fonts, colors</span></span>
<span data-ttu-id="2b52f-p103">O Fabric Core cont?m os elementos principais da linguagem de design, como ?cones, cores, tipo e grade. O Fabric Core ? independente de estrutura. Tanto o Fabric JS como o Fabric React usam o Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="2b52f-p103">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid. Fabric core is framework independent. Both Fabric React and Fabric JS use Fabric Core.</span></span>

<span data-ttu-id="2b52f-113">Para come?ar a usar o Fabric Core:</span><span class="sxs-lookup"><span data-stu-id="2b52f-113">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="2b52f-114">Adicione a refer?ncia da CDN ao HTML da sua p?gina.</span><span class="sxs-lookup"><span data-stu-id="2b52f-114">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. <span data-ttu-id="2b52f-115">Use ?cones e fontes do Fabric.</span><span class="sxs-lookup"><span data-stu-id="2b52f-115">Use Fabric icons and fonts.</span></span> 

    <span data-ttu-id="2b52f-p104">Para usar um ?cone do Fabric, inclua o elemento "i" na sua p?gina e, em seguida, fa?a refer?ncia ?s classes apropriadas. Para controlar o tamanho do ?cone, voc? pode alterar o tamanho da fonte. Por exemplo, o c?digo a seguir mostra como criar um ?cone de tabela muito grande que usa a cor themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="2b52f-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span> 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="2b52f-p105">Para localizar mais ?cones dispon?veis no Office UI Fabric, use o recurso de pesquisa na p?gina [?cones](https://dev.office.com/fabric#/styles/icons). Quando encontrar um ?cone para usar no suplemento, n?o deixe de adicionar um prefixo ao nome do ?cone com `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="2b52f-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://dev.office.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span> 

    <span data-ttu-id="2b52f-121">Para saber mais sobre os tamanhos de fonte e as cores dispon?veis no Office UI Fabric, confira [Tipografia](https://dev.office.com/fabric#/styles/typography) e [Cores](https://dev.office.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="2b52f-121">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://dev.office.com/fabric#/styles/typography) and [Colors](https://dev.office.com/fabric#/styles/colors).</span></span>
 
## <a name="use-fabric-components"></a><span data-ttu-id="2b52f-122">Uso dos componentes do Fabric</span><span class="sxs-lookup"><span data-stu-id="2b52f-122">Use Fabric Components</span></span> 
<span data-ttu-id="2b52f-123">O Fabric oferece uma variedade de componentes da experi?ncia do usu?rio que voc? pode usar para criar o suplemento. Alguns desses componentes incluem:</span><span class="sxs-lookup"><span data-stu-id="2b52f-123">Fabric provides a variety of UX components that you can use to build your add-in, including the following types of components:</span></span>

- <span data-ttu-id="2b52f-124">Componentes de entrada ? por exemplo, bot?o, caixa de sele??o e altern?ncia</span><span class="sxs-lookup"><span data-stu-id="2b52f-124">Input components - for example, Button, Checkbox, and Toggle</span></span>
- <span data-ttu-id="2b52f-125">Componentes de navega??o ? por exemplo, din?mico e trilha</span><span class="sxs-lookup"><span data-stu-id="2b52f-125">Navigation components - for example, Pivot Breadcrumb</span></span>
- <span data-ttu-id="2b52f-126">Componentes de notifica??o ? por exemplo, MessageBar e bal?o</span><span class="sxs-lookup"><span data-stu-id="2b52f-126">Notification components - for example, MessageBar and Callout</span></span>  

<span data-ttu-id="2b52f-p106">Nem todos os componentes do Fabric s?o recomendados para uso em suplementos. Fornecemos diretrizes sobre como usar os componentes recomendados nesta se??o. Por exemplo, para ver orienta??es de como usar um bot?o do Fabric no suplemento, confira [Bot?o](button.md).</span><span class="sxs-lookup"><span data-stu-id="2b52f-p106">Not all Fabric components are recommended for use in add-ins. We provide guidance for how you can use the recommended components in this section. For example, for guidance for using a Fabric button in your add-in, see [Button](button.md).</span></span> 

<span data-ttu-id="2b52f-p107">Voc? pode usar diferentes estruturas do JavaScript, como Angular ou React, para criar o suplemento. Para come?ar a usar componentes do Fabric com sua estrutura, confira os recursos a seguir.</span><span class="sxs-lookup"><span data-stu-id="2b52f-p107">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="2b52f-131">**Estrutura**</span><span class="sxs-lookup"><span data-stu-id="2b52f-131">**Framework**</span></span>|<span data-ttu-id="2b52f-132">**Exemplo**</span><span class="sxs-lookup"><span data-stu-id="2b52f-132">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="2b52f-133">**Rea??o**</span><span class="sxs-lookup"><span data-stu-id="2b52f-133">**React**</span></span>|[<span data-ttu-id="2b52f-134">Uso do Office UI Fabric React em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="2b52f-134">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="2b52f-135">**Angular**</span><span class="sxs-lookup"><span data-stu-id="2b52f-135">**Angular**</span></span>| <span data-ttu-id="2b52f-136">Confira [ngOfficeUIFabric](http://ngofficeuifabric.com/), que ? um projeto comunit?rio com diretivas do Angular 1.5, e [Considere a possibilidade de dispor componentes do Fabric com componentes do Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span><span class="sxs-lookup"><span data-stu-id="2b52f-136">See [ngOfficeUIFabric](http://ngofficeuifabric.com/), which is a community project with Angular 1.5 directives, and [Consider wrapping Fabric components with Angular 2 components](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span></span>|
