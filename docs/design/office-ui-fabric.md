---
title: Office UI Fabric em suplementos do Office
description: Obter uma visão geral de como usar os componentes do Office UI Fabric em Complementos do Office.
ms.date: 2/09/2021
localization_priority: Normal
ms.openlocfilehash: 9799d98d795486203e4bcc23bffc043c2ead6e28
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237676"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="45625-103">Office UI Fabric em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="45625-103">Office UI Fabric in Office Add-ins</span></span>

<span data-ttu-id="45625-104">O Office UI Fabric é uma estrutura de front-end JavaScript para criar experiências de usuário para o Office.</span><span class="sxs-lookup"><span data-stu-id="45625-104">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office.</span></span> <span data-ttu-id="45625-105">O Fabric fornece componentes com foco em efeitos visuais que você pode estender, reformular e usar no suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="45625-105">Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in.</span></span> <span data-ttu-id="45625-106">Como o Fabric usa a linguagem de design da Microsoft, os componentes da experiência de usuário do Fabric são semelhantes a uma extensão natural do Office.</span><span class="sxs-lookup"><span data-stu-id="45625-106">Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span>

<span data-ttu-id="45625-p102">Se estiver criando um suplemento, recomendamos usar o Office UI Fabric para criar a experiência de usuário. O uso do Office UI Fabric é opcional.</span><span class="sxs-lookup"><span data-stu-id="45625-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="45625-109">As seções a seguir explicam como começar a usar o Fabric para atender às suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="45625-109">The following sections explain how to get started using Fabric to meet your requirements.</span></span>

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="45625-110">Uso do Fabric Core: ícones, fontes, cores</span><span class="sxs-lookup"><span data-stu-id="45625-110">Use Fabric Core: icons, fonts, colors</span></span>

<span data-ttu-id="45625-111">O Fabric Core contém os elementos principais da linguagem de design, como ícones, cores, tipo e grade.</span><span class="sxs-lookup"><span data-stu-id="45625-111">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid.</span></span> <span data-ttu-id="45625-112"> O Fabric Core é independente de estrutura.</span><span class="sxs-lookup"><span data-stu-id="45625-112">Fabric core is framework independent.</span></span> <span data-ttu-id="45625-113">O Fabric Core é usado pelo Fabric React e incluído nele.</span><span class="sxs-lookup"><span data-stu-id="45625-113">Fabric Core is used by and included with Fabric React.</span></span>

<span data-ttu-id="45625-114">Para começar a usar o Fabric Core:</span><span class="sxs-lookup"><span data-stu-id="45625-114">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="45625-115">Adicione a referência da CDN ao HTML da sua página.</span><span class="sxs-lookup"><span data-stu-id="45625-115">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. <span data-ttu-id="45625-116">Use ícones e fontes do Fabric.</span><span class="sxs-lookup"><span data-stu-id="45625-116">Use Fabric icons and fonts.</span></span>

    <span data-ttu-id="45625-p104">Para usar um ícone do Fabric, inclua o elemento "i" na sua página e, em seguida, faça referência às classes apropriadas. Para controlar o tamanho do ícone, você pode alterar o tamanho da fonte. Por exemplo, o código a seguir mostra como criar um ícone de tabela muito grande que usa a cor themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="45625-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span>

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="45625-p105">Para localizar mais ícones disponíveis no Office UI Fabric, use o recurso de pesquisa na página [Ícones](https://developer.microsoft.com/fabric#/styles/icons). Quando encontrar um ícone para usar no suplemento, não deixe de adicionar um prefixo ao nome do ícone com `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="45625-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://developer.microsoft.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span>

    <span data-ttu-id="45625-122">Para saber mais sobre os tamanhos de fonte e as cores disponíveis no Office UI Fabric, confira [Tipografia](https://developer.microsoft.com/fabric#/styles/typography) e [Cores](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="45625-122">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://developer.microsoft.com/fabric#/styles/typography) and [Colors](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

## <a name="use-fabric-components"></a><span data-ttu-id="45625-123">Uso dos componentes do Fabric</span><span class="sxs-lookup"><span data-stu-id="45625-123">Use Fabric Components</span></span>

<span data-ttu-id="45625-124">O Fabric fornece uma variedade de componentes da UX que você pode usar para criar seu complemento.</span><span class="sxs-lookup"><span data-stu-id="45625-124">Fabric provides a variety of UX components that you can use to build your add-in.</span></span> <span data-ttu-id="45625-125">Não esperamos que todos os componentes do Fabric sejam usados por um único complemento.</span><span class="sxs-lookup"><span data-stu-id="45625-125">We do not expect that all fabric components will be used by a single add-in.</span></span> <span data-ttu-id="45625-126">Determine os melhores componentes para seu cenário e experiência do usuário [](https://developer.microsoft.com/fabric#/components/breadcrumb) (por exemplo, pode ser difícil exibir corretamente uma navegação de navegação no painel de tarefas).</span><span class="sxs-lookup"><span data-stu-id="45625-126">Determine the best components for your scenario and user experience (for example, it may be hard to properly display a [Breadcrumb](https://developer.microsoft.com/fabric#/components/breadcrumb) in the task pane).</span></span>

<span data-ttu-id="45625-127">Veja a seguir uma lista de componentes comuns da experiência de usuário do [Fabric React](https://developer.microsoft.com/fluentui#/controls/web) que recomendamos para uso em um complemento:</span><span class="sxs-lookup"><span data-stu-id="45625-127">The following is a list of common [Fabric React UX components](https://developer.microsoft.com/fluentui#/controls/web) that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="45625-128">Botão</span><span class="sxs-lookup"><span data-stu-id="45625-128">Button</span></span>](https://developer.microsoft.com/fabric#/components/button)
- [<span data-ttu-id="45625-129">Caixa de seleção</span><span class="sxs-lookup"><span data-stu-id="45625-129">Checkbox</span></span>](https://developer.microsoft.com/fabric#/components/checkbox)
- [<span data-ttu-id="45625-130">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="45625-130">ChoiceGroup</span></span>](https://developer.microsoft.com/fabric#/components/choicegroup)
- [<span data-ttu-id="45625-131">Lista suspensa</span><span class="sxs-lookup"><span data-stu-id="45625-131">Dropdown</span></span>](https://developer.microsoft.com/fabric#/components/dropdown)
- [<span data-ttu-id="45625-132">Rótulo</span><span class="sxs-lookup"><span data-stu-id="45625-132">Label</span></span>](https://developer.microsoft.com/fabric#/components/label)
- [<span data-ttu-id="45625-133">Lista</span><span class="sxs-lookup"><span data-stu-id="45625-133">List</span></span>](https://developer.microsoft.com/fabric#/components/list)
- [<span data-ttu-id="45625-134">Tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="45625-134">Pivot</span></span>](https://developer.microsoft.com/fabric#/components/pivot)
- [<span data-ttu-id="45625-135">Campo de texto</span><span class="sxs-lookup"><span data-stu-id="45625-135">TextField</span></span>](https://developer.microsoft.com/fabric#/components/textfield)
- [<span data-ttu-id="45625-136">Alternância</span><span class="sxs-lookup"><span data-stu-id="45625-136">Toggle</span></span>](https://developer.microsoft.com/fabric#/components/toggle)

<span data-ttu-id="45625-p107">Você pode usar diferentes estruturas do JavaScript, como Angular ou React, para criar o suplemento. Para começar a usar componentes do Fabric com sua estrutura, confira os recursos a seguir.</span><span class="sxs-lookup"><span data-stu-id="45625-p107">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="45625-139">**Framework**</span><span class="sxs-lookup"><span data-stu-id="45625-139">**Framework**</span></span>|<span data-ttu-id="45625-140">**Exemplo**</span><span class="sxs-lookup"><span data-stu-id="45625-140">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="45625-141">**React**</span><span class="sxs-lookup"><span data-stu-id="45625-141">**React**</span></span>|[<span data-ttu-id="45625-142">Uso do Office UI Fabric React em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="45625-142">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="45625-143">**Angular**</span><span class="sxs-lookup"><span data-stu-id="45625-143">**Angular**</span></span>| [<span data-ttu-id="45625-144">Considere a possibilidade de quebra de componentes do Fabric com componentes do Angular 2</span><span class="sxs-lookup"><span data-stu-id="45625-144">Consider wrapping Fabric components with Angular 2 components</span></span>](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
