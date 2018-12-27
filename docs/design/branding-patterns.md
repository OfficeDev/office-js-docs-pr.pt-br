---
title: Diretrizes de design de padrões de identidade visual para suplementos do Office
description: ''
ms.date: 06/26/2018
ms.openlocfilehash: a94e723b222dfe1b004d8b558da59804faf51e69
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433692"
---
# <a name="branding-patterns"></a><span data-ttu-id="e486c-102">Padrões de identidade visual</span><span class="sxs-lookup"><span data-stu-id="e486c-102">Branding patterns</span></span>

<span data-ttu-id="e486c-103">Esses padrões proporcionam visibilidade à marca e contexto aos seus usuários.</span><span class="sxs-lookup"><span data-stu-id="e486c-103">These patterns provide brand visibilty and context to your add-in users.</span></span> 

## <a name="best-practices"></a><span data-ttu-id="e486c-104">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="e486c-104">Best practices</span></span>

|<span data-ttu-id="e486c-105">Fazer</span><span class="sxs-lookup"><span data-stu-id="e486c-105">Do</span></span> |<span data-ttu-id="e486c-106">Não fazer</span><span class="sxs-lookup"><span data-stu-id="e486c-106">Don't</span></span>|
|:---- |:----|
| <span data-ttu-id="e486c-107">Use os componentes familiares de interface do usuário com a aplicação de destaques de identidade visual, tais como tipografia e cor.</span><span class="sxs-lookup"><span data-stu-id="e486c-107">Use familiar UI components with applied branding accents like typography and color.</span></span> | <span data-ttu-id="e486c-108">Não crie novos componentes para a interface do usuário que contradigam a interface de usuário estabelecida do Office.</span><span class="sxs-lookup"><span data-stu-id="e486c-108">Don't invent new UI components that contradict established Office UI.</span></span> | 
| <span data-ttu-id="e486c-109">Aplique a identidade visual de suplemento no rodapé da barra da marca na parte inferior da sua interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="e486c-109">Place your add-in branding in a brand bar footer at the bottom of your UI.</span></span> | <span data-ttu-id="e486c-110">Não repita o nome do painel de tarefas na barra de marca imediatamente adjacente à parte superior da sua interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="e486c-110">Don't repeat your task pane name in an immediately adjacent brand bar at the top of your UI.</span></span> |
| <span data-ttu-id="e486c-111">Use os elementos de marca com moderação.</span><span class="sxs-lookup"><span data-stu-id="e486c-111">Use brand elements sparingly.</span></span> <span data-ttu-id="e486c-112">Ajuste sua solução para o Office de forma complementar.</span><span class="sxs-lookup"><span data-stu-id="e486c-112">Fit your solution into Office such that is complementary.</span></span> | <span data-ttu-id="e486c-113">Não insira elementos da marca de forma excessiva na interface do usuário do Office porque podem distrair e confundir os clientes.</span><span class="sxs-lookup"><span data-stu-id="e486c-113">Don't insert excessively branded elements into Office UI that distract and confuse customers.</span></span> |
| <span data-ttu-id="e486c-114">Verifique se a sua solução é reconhecível e conecta as telas com elementos visuais consistentes.</span><span class="sxs-lookup"><span data-stu-id="e486c-114">Make your solution recognizable and connect your screens together with consistent visual elements.</span></span> | <span data-ttu-id="e486c-115">Não oculte sua solução com elementos visuais aplicados de modo inconsistente e irreconhecíveis.</span><span class="sxs-lookup"><span data-stu-id="e486c-115">Don't hide your solution with unrecognizable and inconsistently applied visual elements.</span></span> |
| <span data-ttu-id="e486c-116">Crie a conexão com um serviço ou negócio relacionado para garantir que os clientes reconheçam e confiem na sua solução.</span><span class="sxs-lookup"><span data-stu-id="e486c-116">Build connection with a parent service or business to ensure that customers know and trust your solution.</span></span> | <span data-ttu-id="e486c-117">Não force os clientes a aprender um novo conceito de marca se já houver um relacionamento útil e compreensível que possa ser aproveitado para criar confiança e valor.</span><span class="sxs-lookup"><span data-stu-id="e486c-117">Don't make customers learn a new brand concept if there is a useful and understandable relationship that can be leveraged to build trust and value.</span></span> |


<span data-ttu-id="e486c-118">Aplique os padrões e componentes a seguir, quando possível, para permitir que os usuários adotem a utilização total do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="e486c-118">Apply the following patterns and components as applicable to allow users to embrace the full utility of your add-in.</span></span>


## <a name="brand-bar"></a><span data-ttu-id="e486c-119">Barra da marca</span><span class="sxs-lookup"><span data-stu-id="e486c-119">Brand bar</span></span>

<span data-ttu-id="e486c-120">A barra da marca é um espaço no rodapé para incluir o nome e o logotipo da marca.</span><span class="sxs-lookup"><span data-stu-id="e486c-120">The brand bar is a space in the footer to include your brand name and logo.</span></span> <span data-ttu-id="e486c-121">Também funciona como um link para o site da marca e um local de acesso opcional.</span><span class="sxs-lookup"><span data-stu-id="e486c-121">It also serves as a link to your brand's website and an optional access location.</span></span>

![Barra de marca – especificações do painel de tarefas da área de trabalho](../images/add-in-brand-bar.png)

## <a name="splash-screen"></a><span data-ttu-id="e486c-123">Tela inicial</span><span class="sxs-lookup"><span data-stu-id="e486c-123">Splash Screen</span></span>

<span data-ttu-id="e486c-124">Use esta tela para exibir a sua identidade visual enquanto o suplemento estiver carregando ou na transição entre estados de interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="e486c-124">Use this screen to display your branding while the add-in is loading or transitioning between UI states.</span></span>

![Tela inicial da marca – especificações do painel de tarefas da área de trabalho](../images/add-in-splash-screen.png)