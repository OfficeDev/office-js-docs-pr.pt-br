---
title: Diretrizes de design de padrões de identidade visual para suplementos do Office
description: Saiba como fazer a marca do seu Office Add-in enquanto permanece compatível com o design visual do Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: b42d3a722e4f8805e8c03d2e1a5db528a66f1202
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076368"
---
# <a name="branding-patterns"></a><span data-ttu-id="181ae-103">Padrões de identidade visual</span><span class="sxs-lookup"><span data-stu-id="181ae-103">Branding patterns</span></span>

<span data-ttu-id="181ae-104">Esses padrões fornecem visibilidade da marca e contexto aos usuários do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="181ae-104">These patterns provide brand visibility and context to your add-in users.</span></span>

## <a name="best-practices"></a><span data-ttu-id="181ae-105">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="181ae-105">Best practices</span></span>

|<span data-ttu-id="181ae-106">Fazer</span><span class="sxs-lookup"><span data-stu-id="181ae-106">Do</span></span> |<span data-ttu-id="181ae-107">Não fazer</span><span class="sxs-lookup"><span data-stu-id="181ae-107">Don't</span></span>|
|:---- |:----|
| <span data-ttu-id="181ae-108">Use os componentes familiares de interface do usuário com a aplicação de destaques de identidade visual, tais como tipografia e cor.</span><span class="sxs-lookup"><span data-stu-id="181ae-108">Use familiar UI components with applied branding accents like typography and color.</span></span> | <span data-ttu-id="181ae-109">Não crie novos componentes para a interface do usuário que contradigam a interface de usuário estabelecida do Office.</span><span class="sxs-lookup"><span data-stu-id="181ae-109">Don't invent new UI components that contradict established Office UI.</span></span> |
| <span data-ttu-id="181ae-110">Aplique a identidade visual de suplemento no rodapé da barra da marca na parte inferior da sua interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="181ae-110">Place your add-in branding in a brand bar footer at the bottom of your UI.</span></span> | <span data-ttu-id="181ae-111">Não repita o nome do painel de tarefas na barra de marca imediatamente adjacente à parte superior da sua interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="181ae-111">Don't repeat your task pane name in an immediately adjacent brand bar at the top of your UI.</span></span> |
| <span data-ttu-id="181ae-112">Use os elementos de marca com moderação.</span><span class="sxs-lookup"><span data-stu-id="181ae-112">Use brand elements sparingly.</span></span> <span data-ttu-id="181ae-113">Ajuste sua solução para o Office de forma complementar.</span><span class="sxs-lookup"><span data-stu-id="181ae-113">Fit your solution into Office such that is complementary.</span></span> | <span data-ttu-id="181ae-114">Não insira elementos da marca de forma excessiva na interface do usuário do Office porque podem distrair e confundir os clientes.</span><span class="sxs-lookup"><span data-stu-id="181ae-114">Don't insert excessively branded elements into Office UI that distract and confuse customers.</span></span> |
| <span data-ttu-id="181ae-115">Verifique se a sua solução é reconhecível e conecta as telas com elementos visuais consistentes.</span><span class="sxs-lookup"><span data-stu-id="181ae-115">Make your solution recognizable and connect your screens together with consistent visual elements.</span></span> | <span data-ttu-id="181ae-116">Não oculte sua solução com elementos visuais aplicados de modo inconsistente e irreconhecíveis.</span><span class="sxs-lookup"><span data-stu-id="181ae-116">Don't hide your solution with unrecognizable and inconsistently applied visual elements.</span></span> |
| <span data-ttu-id="181ae-117">Crie a conexão com um serviço ou negócio relacionado para garantir que os clientes reconheçam e confiem na sua solução.</span><span class="sxs-lookup"><span data-stu-id="181ae-117">Build connection with a parent service or business to ensure that customers know and trust your solution.</span></span> | <span data-ttu-id="181ae-118">Não force os clientes a aprender um novo conceito de marca se já houver um relacionamento útil e compreensível que possa ser aproveitado para criar confiança e valor.</span><span class="sxs-lookup"><span data-stu-id="181ae-118">Don't make customers learn a new brand concept if there is a useful and understandable relationship that can be leveraged to build trust and value.</span></span> |

<span data-ttu-id="181ae-119">Aplique os padrões e componentes a seguir, quando possível, para permitir que os usuários adotem a utilização total do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="181ae-119">Apply the following patterns and components as applicable to allow users to embrace the full utility of your add-in.</span></span>

## <a name="brand-bar"></a><span data-ttu-id="181ae-120">Barra da marca</span><span class="sxs-lookup"><span data-stu-id="181ae-120">Brand Bar</span></span>

<span data-ttu-id="181ae-121">A barra da marca é um espaço no rodapé para incluir o nome e o logotipo da marca.</span><span class="sxs-lookup"><span data-stu-id="181ae-121">The brand bar is a space in the footer to include your brand name and logo.</span></span> <span data-ttu-id="181ae-122">Também funciona como um link para o site da marca e um local de acesso opcional.</span><span class="sxs-lookup"><span data-stu-id="181ae-122">It also serves as a link to your brand's website and an optional access location.</span></span>

![Barra de marcas exibida em um painel de tarefas de um Office de área de trabalho.](../images/add-in-brand-bar.png)

## <a name="splash-screen"></a><span data-ttu-id="181ae-124">Tela inicial</span><span class="sxs-lookup"><span data-stu-id="181ae-124">Splash Screen</span></span>

<span data-ttu-id="181ae-125">Use esta tela para exibir a sua identidade visual enquanto o suplemento estiver carregando ou na transição entre estados de interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="181ae-125">Use this screen to display your branding while the add-in is loading or transitioning between UI states.</span></span>

![Tela inicial da marca exibida em um painel de tarefas de um Office de área de trabalho.](../images/add-in-splash-screen.png)
