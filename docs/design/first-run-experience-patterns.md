---
title: Padrões de tela de apresentação para suplemento dos Office
description: Saiba as práticas recomendadas para projetar experiências de tela de apresentação em suplementos do Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 00785df2cfd2f41b41917ea720c154e24b72f779
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132064"
---
# <a name="first-run-experience-patterns"></a><span data-ttu-id="55afe-103">Padrões de tela de apresentação</span><span class="sxs-lookup"><span data-stu-id="55afe-103">First-run experience patterns</span></span>

<span data-ttu-id="55afe-104">Uma tela de apresentação (FRE) é a introdução de um usuário para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="55afe-104">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="55afe-105">Um FRE é exibida quando um usuário abre um suplemento pela primeira vez e fornece informações sobre as funções, recursos e/ou os benefícios do suplemento.</span><span class="sxs-lookup"><span data-stu-id="55afe-105">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="55afe-106">Essa experiência ajuda a moldar a impressão do usuário de um suplemento e pode influenciar fortemente sua probabilidade de voltar e continuar usando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="55afe-106">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="55afe-107">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="55afe-107">Best practices</span></span>

<span data-ttu-id="55afe-108">Siga estas práticas recomendadas ao criar sua tela de apresentação:</span><span class="sxs-lookup"><span data-stu-id="55afe-108">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="55afe-109">Fazer</span><span class="sxs-lookup"><span data-stu-id="55afe-109">Do</span></span>|<span data-ttu-id="55afe-110">Não fazer</span><span class="sxs-lookup"><span data-stu-id="55afe-110">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="55afe-111">Forneça uma simples e breve introdução para as principais ações do suplemento.</span><span class="sxs-lookup"><span data-stu-id="55afe-111">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="55afe-112">Não inclua informações e legendas que não sejam relevantes ao guia de introdução.</span><span class="sxs-lookup"><span data-stu-id="55afe-112">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="55afe-113">Forneça aos usuários a oportunidade de concluir uma ação que impactará positivamente o uso do add-in.</span><span class="sxs-lookup"><span data-stu-id="55afe-113">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="55afe-114">Não espere que os usuários aprendam tudo ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="55afe-114">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="55afe-115">Concentre-se na ação que fornece o maior valor.</span><span class="sxs-lookup"><span data-stu-id="55afe-115">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="55afe-116">Crie uma experiência envolvente que os usuários desejem concluir.</span><span class="sxs-lookup"><span data-stu-id="55afe-116">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="55afe-117">Não force os usuários a clicar na experiência da tela de apresentação.</span><span class="sxs-lookup"><span data-stu-id="55afe-117">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="55afe-118">Forneça aos usuários uma opção para ignorar a tela de apresentação.</span><span class="sxs-lookup"><span data-stu-id="55afe-118">Give users an option to bypass the first-run experience.</span></span> |

<span data-ttu-id="55afe-119">Considere se mostrar aos usuários a tela de apresentação uma vez ou periodicamente é importante para seu cenário.</span><span class="sxs-lookup"><span data-stu-id="55afe-119">Consider whether showing users the first-run experience once or periodically is important to your scenario.</span></span> <span data-ttu-id="55afe-120">Por exemplo, se o suplemento for usado apenas periodicamente, os usuários poderão ficar menos familiarizados com seu suplemento e poderão se beneficiar de outra interação com a tela de apresentação.</span><span class="sxs-lookup"><span data-stu-id="55afe-120">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>

<span data-ttu-id="55afe-121">Aplique os seguintes padrões, conforme aplicável, para criar ou aprimorar a tela de apresentação do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="55afe-121">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>

## <a name="carousel"></a><span data-ttu-id="55afe-122">Carrossel</span><span class="sxs-lookup"><span data-stu-id="55afe-122">Carousel</span></span>

<span data-ttu-id="55afe-123">O carrossel apresenta aos usuários uma série de recursos ou página de informações antes que eles comecem a usar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="55afe-123">The carousel takes users through a series of features or informational pages before they start using the add-in.</span></span>

<span data-ttu-id="55afe-124">*Figura 1. Permitir que os usuários avancem ou ignorem as páginas iniciais do fluxo de carrossel*</span><span class="sxs-lookup"><span data-stu-id="55afe-124">*Figure 1. Allow users to advance or skip the beginning pages of the carousel flow*</span></span>

![Ilustração mostrando a etapa 1 de um carrossel na experiência de primeira execução de um painel de tarefas de aplicativo da área de trabalho do Office.](../images/add-in-FRE-step-1.png)

<span data-ttu-id="55afe-127">*Figura 2. Minimizar o número de telas de carrossel apenas para o que é necessário para comunicar efetivamente sua mensagem*</span><span class="sxs-lookup"><span data-stu-id="55afe-127">*Figure 2. Minimize the number of carousel screens to only what is needed to effectively communicate your message*</span></span>

![Ilustração mostrando a etapa 2 de um carrossel na experiência de primeira execução de um painel de tarefas de aplicativo da área de trabalho do Office.](../images/add-in-FRE-step-2.png)

<span data-ttu-id="55afe-130">*Figura 3. Fornecer um plano de ação claro para sair da experiência de primeira execução*</span><span class="sxs-lookup"><span data-stu-id="55afe-130">*Figure 3. Provide a clear call to action to exit the first-run-experience*</span></span>

![Ilustração mostrando a etapa 3 de um carrossel na experiência de primeira execução de um painel de tarefas de aplicativo da área de trabalho do Office.](../images/add-in-FRE-step-3.png)

## <a name="value-placemat"></a><span data-ttu-id="55afe-133">Roteiro de valor</span><span class="sxs-lookup"><span data-stu-id="55afe-133">Value Placemat</span></span>

<span data-ttu-id="55afe-134">O posicionamento do valor informa a proposta de valor do seu suplemento com posicionamento do logotipo, uma proposta de valor claramente definida, destaques ou resumo do recurso e uma chamada para ação.</span><span class="sxs-lookup"><span data-stu-id="55afe-134">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>

<span data-ttu-id="55afe-135">*Figura 4. Um valor roteiro com logotipo, proposta de valor claro, Resumo de recursos e plano de ação*</span><span class="sxs-lookup"><span data-stu-id="55afe-135">*Figure 4. A value placemat with logo, clear value proposition, feature summary, and call-to-action*</span></span>

![Ilustração mostrando um valor roteiro na primeira experiência de execução de um painel de tarefas de aplicativo da área de trabalho do Office.](../images/add-in-FRE-value.png)

### <a name="video-placemat"></a><span data-ttu-id="55afe-138">Roteiro de vídeo</span><span class="sxs-lookup"><span data-stu-id="55afe-138">Video Placemat</span></span>

<span data-ttu-id="55afe-139">O roteiro de vídeo mostra aos usuários um vídeo antes que eles comecem a usar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="55afe-139">The video placemat shows users a video before they start using your add-in.</span></span>

<span data-ttu-id="55afe-140">*Figura 5. Tela de apresentação de roteiro-a tela contém uma imagem estática do vídeo com um botão Play e um botão limpar chamada para ação*</span><span class="sxs-lookup"><span data-stu-id="55afe-140">*Figure 5. First run video placemat - The screen contains a still image from the video with a play button and clear call-to-action button*</span></span>

![Ilustração mostrando um roteiro de vídeo na experiência de primeira execução de um painel de tarefas de aplicativo da área de trabalho do Office](../images/add-in-FRE-video.png)

<span data-ttu-id="55afe-142">*Figura 6. Player de vídeo-os usuários apresentados com um vídeo em uma janela de diálogo*</span><span class="sxs-lookup"><span data-stu-id="55afe-142">*Figure 6. Video player - Users presented with a video within a dialog window*</span></span>

![Ilustração mostrando um vídeo em uma janela de diálogo com um aplicativo de área de trabalho e um painel de tarefas de suplemento do Office em segundo plano](../images/add-in-FRE-video-dialog.png)
