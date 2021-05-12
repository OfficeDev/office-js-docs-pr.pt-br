---
title: Padrões de design da experiência do usuário para suplementos do Office
description: Obter uma visão geral dos padrões de design da interface do usuário para Office de complementos, incluindo padrões de navegação, autenticação, primeira-executar e identidade visual.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 8544b56b85a25d522c95546b42a78fe01a3c2586
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330105"
---
# <a name="ux-design-patterns-for-office-add-ins"></a><span data-ttu-id="b04c9-103">Padrões de design da experiência do usuário para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b04c9-103">UX design patterns for Office Add-ins</span></span>

<span data-ttu-id="b04c9-104">O design da experiência do usuário para os suplementos do Office deve fornecer uma experiência atraente para os usuários do Office e estender a experiência geral do Office, ajustando-se perfeitamente à interface do usuário padrão do Office.</span><span class="sxs-lookup"><span data-stu-id="b04c9-104">Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.</span></span>  

<span data-ttu-id="b04c9-105">Nossos padrões de experiência do usuário são compostos de componentes.</span><span class="sxs-lookup"><span data-stu-id="b04c9-105">Our UX patterns are composed of components.</span></span> <span data-ttu-id="b04c9-106">Os componentes são controles que ajudam os clientes a interagir com os elementos do software ou serviço.</span><span class="sxs-lookup"><span data-stu-id="b04c9-106">Components are controls that help your customers interact with elements of your software or service.</span></span> <span data-ttu-id="b04c9-107">Botões, navegação e menus são exemplos de componentes comuns que geralmente possuem comportamentos e estilos consistentes.</span><span class="sxs-lookup"><span data-stu-id="b04c9-107">Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.</span></span>

<span data-ttu-id="b04c9-108">[Os componentes](using-office-ui-fabric-react.md) React de interface do usuário fluente parecem e se comportam como parte do Office, assim como os componentes neutros da [estrutura do Office UI Fabric JS](fabric-core.md).</span><span class="sxs-lookup"><span data-stu-id="b04c9-108">[Fluent UI React components](using-office-ui-fabric-react.md) look and behave like a part of Office, as do the framework-neutral components of [Office UI Fabric JS](fabric-core.md).</span></span> <span data-ttu-id="b04c9-109">Aproveite qualquer conjunto de componentes para se integrar com Office.</span><span class="sxs-lookup"><span data-stu-id="b04c9-109">Take advantage of either set of components to integrate with Office.</span></span> <span data-ttu-id="b04c9-110">Como alternativa, se o seu complemento tiver seu próprio idioma de componente preexistência, você não precisará descartar.</span><span class="sxs-lookup"><span data-stu-id="b04c9-110">Alternatively, if your add-in has its own preexisting component language, you don't need to discard it.</span></span> <span data-ttu-id="b04c9-111">Procure oportunidades para mantê-lo durante a integração ao Office.</span><span class="sxs-lookup"><span data-stu-id="b04c9-111">Look for opportunities to retain it while integrating with Office.</span></span> <span data-ttu-id="b04c9-112">Considere maneiras de trocar elementos estilísticos, remover conflitos ou adotar estilos e comportamentos que removam a confusão para o usuário.</span><span class="sxs-lookup"><span data-stu-id="b04c9-112">Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.</span></span>

<span data-ttu-id="b04c9-113">Os padrões fornecidos são soluções de práticas recomendadas com base em cenários comuns de clientes e pesquisa de experiência do usuário.</span><span class="sxs-lookup"><span data-stu-id="b04c9-113">The provided patterns are best practice solutions based on common customer scenarios and user experience research.</span></span> <span data-ttu-id="b04c9-114">Eles devem fornecer um ponto de entrada rápido para projetar e desenvolver os complementos, bem como orientações para alcançar o equilíbrio entre os elementos de marca da Microsoft e seus próprios.</span><span class="sxs-lookup"><span data-stu-id="b04c9-114">They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft brand elements and your own.</span></span> <span data-ttu-id="b04c9-115">Fornecer uma experiência de usuário moderna e limpa que equilibra elementos de design da linguagem de design da interface do usuário fluente da Microsoft e a identidade de marca exclusiva do parceiro pode ajudar a aumentar a retenção do usuário e a adoção do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="b04c9-115">Providing a clean, modern user experience that balances design elements from Microsoft's Fluent UI design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.</span></span>

<span data-ttu-id="b04c9-116">Use os modelos padrão de experiência do usuário para:</span><span class="sxs-lookup"><span data-stu-id="b04c9-116">Use the UX pattern templates to:</span></span>

* <span data-ttu-id="b04c9-117">Aplicar soluções a cenários comuns de clientes.</span><span class="sxs-lookup"><span data-stu-id="b04c9-117">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="b04c9-118">Aplicar as práticas recomendadas de design.</span><span class="sxs-lookup"><span data-stu-id="b04c9-118">Apply design best practices.</span></span>
* <span data-ttu-id="b04c9-119">Incorpore [componentes e estilos de interface do usuário](https://developer.microsoft.com/fluentui#/get-started) fluentes.</span><span class="sxs-lookup"><span data-stu-id="b04c9-119">Incorporate [Fluent UI](https://developer.microsoft.com/fluentui#/get-started) components and styles.</span></span>
* <span data-ttu-id="b04c9-120">Criar suplementos que se integram visualmente à interface do usuário padrão do Office.</span><span class="sxs-lookup"><span data-stu-id="b04c9-120">Build add-ins that visually integrate with the default Office UI.</span></span>
* <span data-ttu-id="b04c9-121">Idealizar e visualizar a experiência do usuário.</span><span class="sxs-lookup"><span data-stu-id="b04c9-121">Ideate and visualize UX.</span></span>

## <a name="getting-started"></a><span data-ttu-id="b04c9-122">Introdução</span><span class="sxs-lookup"><span data-stu-id="b04c9-122">Getting started</span></span>

<span data-ttu-id="b04c9-123">Os padrões são organizados por ações principais ou experiências comuns em um suplemento.</span><span class="sxs-lookup"><span data-stu-id="b04c9-123">The patterns are organized by key actions or experiences that are common in an add-in.</span></span> <span data-ttu-id="b04c9-124">Os principais grupos são:</span><span class="sxs-lookup"><span data-stu-id="b04c9-124">The main groups are:</span></span>

* [<span data-ttu-id="b04c9-125">Tela de apresentação (FRE)</span><span class="sxs-lookup"><span data-stu-id="b04c9-125">First run experience (FRE)</span></span>](../design/first-run-experience-patterns.md)
* [<span data-ttu-id="b04c9-126">Autenticação</span><span class="sxs-lookup"><span data-stu-id="b04c9-126">Authentication</span></span>](../design/authentication-patterns.md)
* [<span data-ttu-id="b04c9-127">Navegação</span><span class="sxs-lookup"><span data-stu-id="b04c9-127">Navigation</span></span>](../design/navigation-patterns.md)
* [<span data-ttu-id="b04c9-128">Design de identidade Visual</span><span class="sxs-lookup"><span data-stu-id="b04c9-128">Branding Design</span></span>](../design/branding-patterns.md)

<span data-ttu-id="b04c9-129">Navegar por cada agrupamento para ter uma ideia de como você pode projetar o suplemento usando as práticas recomendadas.</span><span class="sxs-lookup"><span data-stu-id="b04c9-129">Browse each grouping to get an idea of how you can design your add-in using best practices.</span></span>

> [!NOTE]
> <span data-ttu-id="b04c9-130">As telas de exemplo mostradas ao longo desta documentação, estão projetadas e exibidas na resolução de **1366x768**.</span><span class="sxs-lookup"><span data-stu-id="b04c9-130">The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**.</span></span>

## <a name="see-also"></a><span data-ttu-id="b04c9-131">Confira também</span><span class="sxs-lookup"><span data-stu-id="b04c9-131">See also</span></span>

* [<span data-ttu-id="b04c9-132">Kits de ferramentas de design</span><span class="sxs-lookup"><span data-stu-id="b04c9-132">Design tool kits</span></span>](design-toolkits.md)
* [<span data-ttu-id="b04c9-133">Interface do usuário do Fluent</span><span class="sxs-lookup"><span data-stu-id="b04c9-133">Fluent UI</span></span>](https://developer.microsoft.com/fluentui#)
* [<span data-ttu-id="b04c9-134">Práticas recomendadas para o desenvolvimento de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b04c9-134">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
* [<span data-ttu-id="b04c9-135">Interface do usuário do Fluent React em Office de complementos</span><span class="sxs-lookup"><span data-stu-id="b04c9-135">Fluent UI React in Office Add-ins</span></span>](using-office-ui-fabric-react.md)
