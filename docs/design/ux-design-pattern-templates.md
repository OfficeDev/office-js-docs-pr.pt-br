---
title: Padrões de design da experiência do usuário para suplementos do Office
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 40b36fb138169bdf848e5f58569e6fc3dee8c09b
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871539"
---
# <a name="ux-design-patterns-for-office-add-ins"></a><span data-ttu-id="58c26-102">Padrões de design da experiência do usuário para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="58c26-102">UX design patterns for Office Add-ins</span></span>

<span data-ttu-id="58c26-103">O design da experiência do usuário para os suplementos do Office deve fornecer uma experiência atraente para os usuários do Office e estender a experiência geral do Office, ajustando-se perfeitamente à interface do usuário padrão do Office.</span><span class="sxs-lookup"><span data-stu-id="58c26-103">Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.</span></span>  

<span data-ttu-id="58c26-104">Nossos padrões de experiência do usuário são compostos de componentes.</span><span class="sxs-lookup"><span data-stu-id="58c26-104">Our UX patterns are composed of components.</span></span> <span data-ttu-id="58c26-105">Os componentes são controles que ajudam os clientes a interagir com os elementos do software ou serviço.</span><span class="sxs-lookup"><span data-stu-id="58c26-105">Components are controls that help your customers interact with elements of your software or service.</span></span> <span data-ttu-id="58c26-106">Botões, navegação e menus são exemplos de componentes comuns que geralmente possuem comportamentos e estilos consistentes.</span><span class="sxs-lookup"><span data-stu-id="58c26-106">Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.</span></span>

<span data-ttu-id="58c26-107">O Office UI Fabric renderiza componentes que têm aparência e comportamento como os de uma parte do Office.</span><span class="sxs-lookup"><span data-stu-id="58c26-107">Office UI Fabric renders components that look and behave like a part of Office.</span></span> <span data-ttu-id="58c26-108">Aproveite o Fabric para se integrar facilmente ao Office.</span><span class="sxs-lookup"><span data-stu-id="58c26-108">Take advantage of Fabric to easily integrate with Office.</span></span> <span data-ttu-id="58c26-109">Se o suplemento tiver sua própria linguagem de componente pré-existente, não será necessário descartá-lo para usar o Fabric.</span><span class="sxs-lookup"><span data-stu-id="58c26-109">If your add-in has its own preexisting component language, you don't need to discard it in favor of Fabric.</span></span> <span data-ttu-id="58c26-110">Procure oportunidades para mantê-lo durante a integração ao Office.</span><span class="sxs-lookup"><span data-stu-id="58c26-110">Look for opportunities to retain it while integrating with Office.</span></span> <span data-ttu-id="58c26-111">Considere maneiras de trocar elementos estilísticos, remover conflitos ou adotar estilos e comportamentos que removam a confusão para o usuário.</span><span class="sxs-lookup"><span data-stu-id="58c26-111">Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.</span></span>

<span data-ttu-id="58c26-112">Os padrões fornecidos são soluções de práticas recomendadas com base em cenários comuns de clientes e pesquisa de experiência do usuário.</span><span class="sxs-lookup"><span data-stu-id="58c26-112">The provided patterns are best practice solutions based on common customer scenarios and user experience research.</span></span> <span data-ttu-id="58c26-113">Eles servem para fornecer um ponto de entrada rápido para projetar e desenvolver suplementos, bem como orientação para alcançar o equilíbrio entre os elementos da Microsoft e da marca.</span><span class="sxs-lookup"><span data-stu-id="58c26-113">They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft and brand elements.</span></span> <span data-ttu-id="58c26-114">Proporcionar uma experiência de usuário limpa e moderna que equilibre elementos de design da linguagem de design do Microsoft Fabric e a identidade de marca exclusiva do parceiro pode ajudar a aumentar a retenção de usuários e a adoção do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="58c26-114">Providing a clean, modern user experience that balances design elements from Microsoft's Fabric design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.</span></span>

<span data-ttu-id="58c26-115">Use os modelos padrão de experiência do usuário para:</span><span class="sxs-lookup"><span data-stu-id="58c26-115">Use the UX pattern templates to:</span></span>

* <span data-ttu-id="58c26-116">Aplicar soluções a cenários comuns de clientes.</span><span class="sxs-lookup"><span data-stu-id="58c26-116">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="58c26-117">Aplicar as práticas recomendadas de design.</span><span class="sxs-lookup"><span data-stu-id="58c26-117">Apply design best practices.</span></span>
* <span data-ttu-id="58c26-118">Incorporar componentes e estilos do [Office UI Fabric](https://developer.microsoft.com/fabric#/get-started).</span><span class="sxs-lookup"><span data-stu-id="58c26-118">Incorporate [Office UI Fabric](https://developer.microsoft.com/fabric#/get-started) components and styles.</span></span>
* <span data-ttu-id="58c26-119">Criar suplementos que se integram visualmente à interface do usuário padrão do Office.</span><span class="sxs-lookup"><span data-stu-id="58c26-119">Build add-ins that visually integrate with the default Office UI.</span></span>
* <span data-ttu-id="58c26-120">Idealizar e visualizar a experiência do usuário.</span><span class="sxs-lookup"><span data-stu-id="58c26-120">Ideate and visualize UX.</span></span>

## <a name="getting-started"></a><span data-ttu-id="58c26-121">Introdução</span><span class="sxs-lookup"><span data-stu-id="58c26-121">Getting started</span></span>

<span data-ttu-id="58c26-122">Os padrões são organizados por ações principais ou experiências comuns em um suplemento.</span><span class="sxs-lookup"><span data-stu-id="58c26-122">The patterns are organized by key actions or experiences that are common in an add-in.</span></span> <span data-ttu-id="58c26-123">Os principais grupos são:</span><span class="sxs-lookup"><span data-stu-id="58c26-123">The main groups are:</span></span>

* [<span data-ttu-id="58c26-124">Tela de apresentação (FRE)</span><span class="sxs-lookup"><span data-stu-id="58c26-124">First run experience (FRE)</span></span>](../design/first-run-experience-patterns.md)
* [<span data-ttu-id="58c26-125">Autenticação</span><span class="sxs-lookup"><span data-stu-id="58c26-125">Authentication</span></span>](../design/authentication-patterns.md)
* [<span data-ttu-id="58c26-126">Navegação</span><span class="sxs-lookup"><span data-stu-id="58c26-126">Navigation</span></span>](../design/navigation-patterns.md)
* [<span data-ttu-id="58c26-127">Design de identidade Visual</span><span class="sxs-lookup"><span data-stu-id="58c26-127">Branding Design</span></span>](../design/branding-patterns.md)

<span data-ttu-id="58c26-128">Navegar por cada agrupamento para ter uma ideia de como você pode projetar o suplemento usando as práticas recomendadas.</span><span class="sxs-lookup"><span data-stu-id="58c26-128">Browse each grouping to get an idea of how you can design your add-in using best practices.</span></span>

> [!NOTE]
> <span data-ttu-id="58c26-129">As telas de exemplo mostradas ao longo desta documentação, estão projetadas e exibidas na resolução de **1366x768**.</span><span class="sxs-lookup"><span data-stu-id="58c26-129">NOTE: The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**</span></span>

## <a name="see-also"></a><span data-ttu-id="58c26-130">Confira também</span><span class="sxs-lookup"><span data-stu-id="58c26-130">See also</span></span>

* [<span data-ttu-id="58c26-131">Kits de ferramentas de design</span><span class="sxs-lookup"><span data-stu-id="58c26-131">Design toolkits</span></span>](design-toolkits.md)
* [<span data-ttu-id="58c26-132">Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="58c26-132">Office UI Fabric</span></span>](https://developer.microsoft.com/fabric)
* [<span data-ttu-id="58c26-133">Práticas recomendadas para o desenvolvimento de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="58c26-133">Best practices for developing Office Add-ins</span></span>](/office/dev/add-ins/concepts/add-in-development-best-practices)
* [<span data-ttu-id="58c26-134">Introdução ao uso do Fabric React</span><span class="sxs-lookup"><span data-stu-id="58c26-134">Get started using Fabric React</span></span>](/office/dev/add-ins/design/using-office-ui-fabric-react)
