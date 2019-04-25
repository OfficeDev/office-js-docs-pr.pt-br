---
title: Diretrizes de design de padrões de identidade visual para suplementos do Office
description: ''
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 6de9962f82a4d07f94ca34cff5ccc3622f80c5d3
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32446990"
---
# <a name="branding-patterns"></a><span data-ttu-id="dd23a-102">Padrões de identidade visual</span><span class="sxs-lookup"><span data-stu-id="dd23a-102">Branding patterns</span></span>

<span data-ttu-id="dd23a-103">Esses padrões proporcionam visibilidade à marca e contexto aos seus usuários.</span><span class="sxs-lookup"><span data-stu-id="dd23a-103">These patterns provide brand visibilty and context to your add-in users.</span></span> 

## <a name="best-practices"></a><span data-ttu-id="dd23a-104">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="dd23a-104">Best practices</span></span>

|<span data-ttu-id="dd23a-105">Fazer</span><span class="sxs-lookup"><span data-stu-id="dd23a-105">Do</span></span> |<span data-ttu-id="dd23a-106">Não fazer</span><span class="sxs-lookup"><span data-stu-id="dd23a-106">Don't</span></span>|
|:---- |:----|
| <span data-ttu-id="dd23a-107">Use os componentes familiares de interface do usuário com a aplicação de destaques de identidade visual, tais como tipografia e cor.</span><span class="sxs-lookup"><span data-stu-id="dd23a-107">Use familiar UI components with applied branding accents like typography and color.</span></span> | <span data-ttu-id="dd23a-108">Não crie novos componentes para a interface do usuário que contradigam a interface de usuário estabelecida do Office.</span><span class="sxs-lookup"><span data-stu-id="dd23a-108">Don't invent new UI components that contradict established Office UI.</span></span> | 
| <span data-ttu-id="dd23a-109">Aplique a identidade visual de suplemento no rodapé da barra da marca na parte inferior da sua interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="dd23a-109">Place your add-in branding in a brand bar footer at the bottom of your UI.</span></span> | <span data-ttu-id="dd23a-110">Não repita o nome do painel de tarefas na barra de marca imediatamente adjacente à parte superior da sua interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="dd23a-110">Don't repeat your task pane name in an immediately adjacent brand bar at the top of your UI.</span></span> |
| <span data-ttu-id="dd23a-111">Use os elementos de marca com moderação.</span><span class="sxs-lookup"><span data-stu-id="dd23a-111">Use brand elements sparingly.</span></span> <span data-ttu-id="dd23a-112">Ajuste sua solução para o Office de forma complementar.</span><span class="sxs-lookup"><span data-stu-id="dd23a-112">Fit your solution into Office such that is complementary.</span></span> | <span data-ttu-id="dd23a-113">Não insira elementos da marca de forma excessiva na interface do usuário do Office porque podem distrair e confundir os clientes.</span><span class="sxs-lookup"><span data-stu-id="dd23a-113">Don't insert excessively branded elements into Office UI that distract and confuse customers.</span></span> |
| <span data-ttu-id="dd23a-114">Verifique se a sua solução é reconhecível e conecta as telas com elementos visuais consistentes.</span><span class="sxs-lookup"><span data-stu-id="dd23a-114">Make your solution recognizable and connect your screens together with consistent visual elements.</span></span> | <span data-ttu-id="dd23a-115">Não oculte sua solução com elementos visuais aplicados de modo inconsistente e irreconhecíveis.</span><span class="sxs-lookup"><span data-stu-id="dd23a-115">Don't hide your solution with unrecognizable and inconsistently applied visual elements.</span></span> |
| <span data-ttu-id="dd23a-116">Crie a conexão com um serviço ou negócio relacionado para garantir que os clientes reconheçam e confiem na sua solução.</span><span class="sxs-lookup"><span data-stu-id="dd23a-116">Build connection with a parent service or business to ensure that customers know and trust your solution.</span></span> | <span data-ttu-id="dd23a-117">Não force os clientes a aprender um novo conceito de marca se já houver um relacionamento útil e compreensível que possa ser aproveitado para criar confiança e valor.</span><span class="sxs-lookup"><span data-stu-id="dd23a-117">Don't make customers learn a new brand concept if there is a useful and understandable relationship that can be leveraged to build trust and value.</span></span> |


<span data-ttu-id="dd23a-118">Aplique os padrões e componentes a seguir, quando possível, para permitir que os usuários adotem a utilização total do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="dd23a-118">Apply the following patterns and components as applicable to allow users to embrace the full utility of your add-in.</span></span>


## <a name="brand-bar"></a><span data-ttu-id="dd23a-119">Barra da marca</span><span class="sxs-lookup"><span data-stu-id="dd23a-119">Brand Bar</span></span>

<span data-ttu-id="dd23a-120">A barra da marca é um espaço no rodapé para incluir o nome e o logotipo da marca.</span><span class="sxs-lookup"><span data-stu-id="dd23a-120">The brand bar is a space in the footer to include your brand name and logo.</span></span> <span data-ttu-id="dd23a-121">Também funciona como um link para o site da marca e um local de acesso opcional.</span><span class="sxs-lookup"><span data-stu-id="dd23a-121">It also serves as a link to your brand's website and an optional access location.</span></span>

![Barra de marca – especificações do painel de tarefas da área de trabalho](../images/add-in-brand-bar.png)

## <a name="splash-screen"></a><span data-ttu-id="dd23a-123">Tela inicial</span><span class="sxs-lookup"><span data-stu-id="dd23a-123">Splash Screen</span></span>

<span data-ttu-id="dd23a-124">Use esta tela para exibir a sua identidade visual enquanto o suplemento estiver carregando ou na transição entre estados de interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="dd23a-124">Use this screen to display your branding while the add-in is loading or transitioning between UI states.</span></span>

![Tela inicial da marca – especificações do painel de tarefas da área de trabalho](../images/add-in-splash-screen.png)
