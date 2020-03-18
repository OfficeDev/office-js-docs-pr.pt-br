---
title: Padrões de tela de apresentação para suplemento dos Office
description: ''
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 85f8e4f7e0082e00ad5064333470f589e449af45
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42688505"
---
# <a name="first-run-experience-patterns"></a><span data-ttu-id="be982-102">Padrões de tela de apresentação</span><span class="sxs-lookup"><span data-stu-id="be982-102">First-run experience patterns</span></span>

<span data-ttu-id="be982-103">Uma tela de apresentação (FRE) é a introdução de um usuário para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="be982-103">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="be982-104">Um FRE é exibida quando um usuário abre um suplemento pela primeira vez e fornece informações sobre as funções, recursos e/ou os benefícios do suplemento.</span><span class="sxs-lookup"><span data-stu-id="be982-104">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="be982-105">Essa experiência ajuda a moldar a impressão do usuário de um suplemento e pode influenciar fortemente sua probabilidade de voltar e continuar usando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="be982-105">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="be982-106">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="be982-106">Best practices</span></span>


<span data-ttu-id="be982-107">Siga estas práticas recomendadas ao criar sua tela de apresentação:</span><span class="sxs-lookup"><span data-stu-id="be982-107">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="be982-108">Fazer</span><span class="sxs-lookup"><span data-stu-id="be982-108">Do</span></span>|<span data-ttu-id="be982-109">Não fazer</span><span class="sxs-lookup"><span data-stu-id="be982-109">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="be982-110">Forneça uma simples e breve introdução para as principais ações do suplemento.</span><span class="sxs-lookup"><span data-stu-id="be982-110">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="be982-111">Não inclua informações e legendas que não sejam relevantes ao guia de introdução.</span><span class="sxs-lookup"><span data-stu-id="be982-111">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="be982-112">Forneça aos usuários a oportunidade de concluir uma ação que impactará positivamente o uso do add-in.</span><span class="sxs-lookup"><span data-stu-id="be982-112">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="be982-113">Não espere que os usuários aprendam tudo ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="be982-113">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="be982-114">Concentre-se na ação que fornece o maior valor.</span><span class="sxs-lookup"><span data-stu-id="be982-114">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="be982-115">Crie uma experiência envolvente que os usuários desejem concluir.</span><span class="sxs-lookup"><span data-stu-id="be982-115">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="be982-116">Não force os usuários a clicar na experiência da tela de apresentação.</span><span class="sxs-lookup"><span data-stu-id="be982-116">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="be982-117">Forneça aos usuários uma opção para ignorar a tela de apresentação.</span><span class="sxs-lookup"><span data-stu-id="be982-117">Give users an option to bypass the first-run experience.</span></span> |



<span data-ttu-id="be982-118">Considere se mostrar aos usuários a tela de apresentação uma vez ou periodicamente é importante para seu cenário.</span><span class="sxs-lookup"><span data-stu-id="be982-118">Consider whether showing users the first-run experience once or periodically is important to your scenario.</span></span> <span data-ttu-id="be982-119">Por exemplo, se o suplemento for usado apenas periodicamente, os usuários poderão ficar menos familiarizados com seu suplemento e poderão se beneficiar de outra interação com a tela de apresentação.</span><span class="sxs-lookup"><span data-stu-id="be982-119">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>



<span data-ttu-id="be982-120">Aplique os seguintes padrões, conforme aplicável, para criar ou aprimorar a tela de apresentação do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="be982-120">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>



## <a name="carousel"></a><span data-ttu-id="be982-121">Carrossel</span><span class="sxs-lookup"><span data-stu-id="be982-121">Carousel</span></span>


<span data-ttu-id="be982-122">O carrossel apresenta aos usuários uma série de recursos ou página de informações antes que eles comecem a usar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="be982-122">The carousel takes users through a series of features or informational pages before they start using the add-in.</span></span>

<span data-ttu-id="be982-123">*Figura 1: Permita que os usuários avancem ou pulem as páginas iniciais do fluxo do carrossel.*
![ Apresentação - carrossel - Especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-step-1.png)</span><span class="sxs-lookup"><span data-stu-id="be982-123">*Figure 1: Allow users to advance or skip the beginning pages of the carousel flow.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-1.png)</span></span>



<span data-ttu-id="be982-124">*Figura 2: Minimize o número de telas do carrossel que você apresenta ao usuário somente para as que são necessárias para comunicar efetivamente sua mensagem*
![ Apresentação - carrossel - Especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-step-2.png)</span><span class="sxs-lookup"><span data-stu-id="be982-124">*Figure 2: Minimize the number of carousel screens you present to the user to only what is needed to effectively communicate your message*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-2.png)</span></span>


<span data-ttu-id="be982-125">*Figura 3: Forneça um apelo à ação claro para sair da tela de apresentação.*
![ Apresentação - carrossel - Especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-step-3.png)</span><span class="sxs-lookup"><span data-stu-id="be982-125">*Figure 3: Provide a clear call to action to exit the first-run-experience.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-3.png)</span></span>



## <a name="value-placemat"></a><span data-ttu-id="be982-126">Roteiro de valor</span><span class="sxs-lookup"><span data-stu-id="be982-126">Value Placemat</span></span>

<span data-ttu-id="be982-127">O posicionamento do valor informa a proposta de valor do seu suplemento com posicionamento do logotipo, uma proposta de valor claramente definida, destaques ou resumo do recurso e uma chamada para ação.</span><span class="sxs-lookup"><span data-stu-id="be982-127">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>



<span data-ttu-id="be982-128">![Apresentação - roteiro de valor - Especificações do painel de tarefas da área de trabalho ](../images/add-in-FRE-value.png)
\* Um roteiro de valor com logotipo, proposição de valor clara, resumo de recurso e chamada para ação.\*</span><span class="sxs-lookup"><span data-stu-id="be982-128">![First Run - Value Placemat - Specifications for desktop task pane](../images/add-in-FRE-value.png)
*A value placemat with logo, clear value proposition, feature summary, and call to action.*</span></span>


### <a name="video-placemat"></a><span data-ttu-id="be982-129">Roteiro de vídeo</span><span class="sxs-lookup"><span data-stu-id="be982-129">Video Placemat</span></span>

<span data-ttu-id="be982-130">O roteiro de vídeo mostra aos usuários um vídeo antes que eles comecem a usar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="be982-130">The video placemat shows users a video before they start using your add-in.</span></span>


<span data-ttu-id="be982-131">\*Figura 1: Apresentação do roteiro - A tela contém uma imagem estática do vídeo com um botão de reprodução e um botão de apelo para ação. \*![Roteiro de vídeo - Especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-video.png)</span><span class="sxs-lookup"><span data-stu-id="be982-131">*Figure 1: First Run Placemat - The screen contains a still image from the video with a play button and clear call to action button.*![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video.png)</span></span>



<span data-ttu-id="be982-132">*Figura 2: Player de vídeo - os usuários são apresentados a um vídeo em uma janela de diálogo.*
![ Apresentação de vídeo - Especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-video-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="be982-132">*Figure 2: Video Player - Users are presented with a video within a dialog window.*
![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video-dialog.png)</span></span>
