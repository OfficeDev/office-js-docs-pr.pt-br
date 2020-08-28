---
title: Padrões de tela de apresentação para suplemento dos Office
description: Saiba as práticas recomendadas para projetar experiências de tela de apresentação em suplementos do Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: c0528e869dd8ee7fe779785fb1a9b6d347deab75
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292950"
---
# <a name="first-run-experience-patterns"></a><span data-ttu-id="7e598-103">Padrões de tela de apresentação</span><span class="sxs-lookup"><span data-stu-id="7e598-103">First-run experience patterns</span></span>

<span data-ttu-id="7e598-104">Uma tela de apresentação (FRE) é a introdução de um usuário para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="7e598-104">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="7e598-105">Um FRE é exibida quando um usuário abre um suplemento pela primeira vez e fornece informações sobre as funções, recursos e/ou os benefícios do suplemento.</span><span class="sxs-lookup"><span data-stu-id="7e598-105">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="7e598-106">Essa experiência ajuda a moldar a impressão do usuário de um suplemento e pode influenciar fortemente sua probabilidade de voltar e continuar usando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="7e598-106">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="7e598-107">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="7e598-107">Best practices</span></span>


<span data-ttu-id="7e598-108">Siga estas práticas recomendadas ao criar sua tela de apresentação:</span><span class="sxs-lookup"><span data-stu-id="7e598-108">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="7e598-109">Fazer</span><span class="sxs-lookup"><span data-stu-id="7e598-109">Do</span></span>|<span data-ttu-id="7e598-110">Não fazer</span><span class="sxs-lookup"><span data-stu-id="7e598-110">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="7e598-111">Forneça uma simples e breve introdução para as principais ações do suplemento.</span><span class="sxs-lookup"><span data-stu-id="7e598-111">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="7e598-112">Não inclua informações e legendas que não sejam relevantes ao guia de introdução.</span><span class="sxs-lookup"><span data-stu-id="7e598-112">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="7e598-113">Forneça aos usuários a oportunidade de concluir uma ação que impactará positivamente o uso do add-in.</span><span class="sxs-lookup"><span data-stu-id="7e598-113">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="7e598-114">Não espere que os usuários aprendam tudo ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="7e598-114">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="7e598-115">Concentre-se na ação que fornece o maior valor.</span><span class="sxs-lookup"><span data-stu-id="7e598-115">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="7e598-116">Crie uma experiência envolvente que os usuários desejem concluir.</span><span class="sxs-lookup"><span data-stu-id="7e598-116">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="7e598-117">Não force os usuários a clicar na experiência da tela de apresentação.</span><span class="sxs-lookup"><span data-stu-id="7e598-117">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="7e598-118">Forneça aos usuários uma opção para ignorar a tela de apresentação.</span><span class="sxs-lookup"><span data-stu-id="7e598-118">Give users an option to bypass the first-run experience.</span></span> |



<span data-ttu-id="7e598-119">Considere se mostrar aos usuários a tela de apresentação uma vez ou periodicamente é importante para seu cenário.</span><span class="sxs-lookup"><span data-stu-id="7e598-119">Consider whether showing users the first-run experience once or periodically is important to your scenario.</span></span> <span data-ttu-id="7e598-120">Por exemplo, se o suplemento for usado apenas periodicamente, os usuários poderão ficar menos familiarizados com seu suplemento e poderão se beneficiar de outra interação com a tela de apresentação.</span><span class="sxs-lookup"><span data-stu-id="7e598-120">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>



<span data-ttu-id="7e598-121">Aplique os seguintes padrões, conforme aplicável, para criar ou aprimorar a tela de apresentação do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="7e598-121">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>



## <a name="carousel"></a><span data-ttu-id="7e598-122">Carrossel</span><span class="sxs-lookup"><span data-stu-id="7e598-122">Carousel</span></span>


<span data-ttu-id="7e598-123">O carrossel apresenta aos usuários uma série de recursos ou página de informações antes que eles comecem a usar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="7e598-123">The carousel takes users through a series of features or informational pages before they start using the add-in.</span></span>

<span data-ttu-id="7e598-124">*Figura 1: permitir que os usuários avancem ou ignorem as páginas iniciais do fluxo de carrossel.* 
 ![ Primeira execução-carrossel etapa 1-especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-step-1.png)</span><span class="sxs-lookup"><span data-stu-id="7e598-124">*Figure 1: Allow users to advance or skip the beginning pages of the carousel flow.*
![First Run - Carousel Step 1 - Specifications for desktop task pane](../images/add-in-FRE-step-1.png)</span></span>



<span data-ttu-id="7e598-125">*Figura 2: minimize o número de telas de carrossel que você apresenta ao usuário apenas para o que é necessário para comunicar efetivamente sua mensagem.* 
 ![ Primeira execução-carrossel etapa 2-especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-step-2.png)</span><span class="sxs-lookup"><span data-stu-id="7e598-125">*Figure 2: Minimize the number of carousel screens you present to the user to only what is needed to effectively communicate your message.*
![First Run - Carousel Step 2 - Specifications for desktop task pane](../images/add-in-FRE-step-2.png)</span></span>


<span data-ttu-id="7e598-126">*Figura 3: forneça um plano de ação claro para sair da experiência de primeira execução.* 
 ![ Primeira execução-carrossel etapa 3-especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-step-3.png)</span><span class="sxs-lookup"><span data-stu-id="7e598-126">*Figure 3: Provide a clear call to action to exit the first-run-experience.*
![First Run - Carousel Step 3 - Specifications for desktop task pane](../images/add-in-FRE-step-3.png)</span></span>



## <a name="value-placemat"></a><span data-ttu-id="7e598-127">Roteiro de valor</span><span class="sxs-lookup"><span data-stu-id="7e598-127">Value Placemat</span></span>

<span data-ttu-id="7e598-128">O posicionamento do valor informa a proposta de valor do seu suplemento com posicionamento do logotipo, uma proposta de valor claramente definida, destaques ou resumo do recurso e uma chamada para ação.</span><span class="sxs-lookup"><span data-stu-id="7e598-128">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>



<span data-ttu-id="7e598-129">![Roteiro de primeiro valor de execução-especificações do painel de tarefas da área de trabalho ](../images/add-in-FRE-value.png)
 *um valor roteiro com logotipo, proposta de valor clara, Resumo de recursos e plano de ação.*</span><span class="sxs-lookup"><span data-stu-id="7e598-129">![First Run - Value Placemat - Specifications for desktop task pane](../images/add-in-FRE-value.png)
*A value placemat with logo, clear value proposition, feature summary, and call-to-action.*</span></span>


### <a name="video-placemat"></a><span data-ttu-id="7e598-130">Roteiro de vídeo</span><span class="sxs-lookup"><span data-stu-id="7e598-130">Video Placemat</span></span>

<span data-ttu-id="7e598-131">O roteiro de vídeo mostra aos usuários um vídeo antes que eles comecem a usar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="7e598-131">The video placemat shows users a video before they start using your add-in.</span></span>


<span data-ttu-id="7e598-132">*Figura 1: primeira execução roteiro-a tela contém uma imagem estática do vídeo com um botão Play e um botão limpar chamada para ação.* 
 ![ Roteiro de vídeo – especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-video.png)</span><span class="sxs-lookup"><span data-stu-id="7e598-132">*Figure 1: First Run Placemat - The screen contains a still image from the video with a play button and clear call-to-action button.*
![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video.png)</span></span>



<span data-ttu-id="7e598-133">*Figura 2: player de vídeo-os usuários são apresentados com um vídeo em uma janela de diálogo.* 
 ![ Vídeo roteiro-diálogo-especificações do painel de tarefas da área de trabalho](../images/add-in-FRE-video-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="7e598-133">*Figure 2: Video Player - Users are presented with a video within a dialog window.*
![Video Placemat - Dialog - Specifications for desktop task pane](../images/add-in-FRE-video-dialog.png)</span></span>
