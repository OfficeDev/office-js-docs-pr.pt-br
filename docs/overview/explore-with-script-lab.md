---
title: Explorar a API JavaScript do Office usando o script Lab
description: Use o script Lab para explorar a API do Office JS e a funcionalidade de protótipo.
ms.topic: article
ms.date: 06/07/2019
localization_priority: Normal
ms.openlocfilehash: 0bab566b08ba25dd3c01cff72f331b2dc9ce304d
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910187"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="5e333-103">Explorar a API JavaScript do Office usando o script Lab</span><span class="sxs-lookup"><span data-stu-id="5e333-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="5e333-104">O [suplemento de laboratório de script](https://store.office.com/app.aspx?assetid=WA104380862), que está disponível gratuitamente na Office Store, permite explorar a API JavaScript do Office enquanto você estiver trabalhando em um programa do Office, como Excel ou Word.</span><span class="sxs-lookup"><span data-stu-id="5e333-104">The [Script Lab add-in](https://store.office.com/app.aspx?assetid=WA104380862), which is available free from the Office store, enables you to explore the Office JavaScript API while you are working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="5e333-105">O script Lab é uma ferramenta conveniente para adicionar ao seu kit de ferramentas de desenvolvimento conforme você protótipo e verificar a funcionalidade desejada no seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="5e333-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="5e333-106">O que é o script Lab?</span><span class="sxs-lookup"><span data-stu-id="5e333-106">What is Script Lab?</span></span>

<span data-ttu-id="5e333-107">O script Lab é uma ferramenta para qualquer pessoa que deseje saber como desenvolver suplementos do Office usando a API JavaScript do Office no Excel, no Word ou no PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="5e333-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="5e333-108">Ele fornece o IntelliSense para que você possa ver o que está disponível e foi criado na estrutura de Mônaco, a mesma estrutura usada pelo Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="5e333-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="5e333-109">Por meio do laboratório de scripts, você pode acessar uma biblioteca de exemplos para experimentar rapidamente recursos ou pode escolher um exemplo como base para seu próprio código.</span><span class="sxs-lookup"><span data-stu-id="5e333-109">Through Script Lab, you can access a library of samples to quickly try out features or you can choose a sample as the base for your own code.</span></span> <span data-ttu-id="5e333-110">Você também é bem-vindo à expansão da biblioteca de amostra adicionando trechos ao [repositório Office-js-Snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).</span><span class="sxs-lookup"><span data-stu-id="5e333-110">You are also welcome to expand the sample library by adding snippets to the [office-js-snippets repo](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).</span></span> <span data-ttu-id="5e333-111">Outro recurso interessante do laboratório de scripts é a funcionalidade beta ou de visualização está disponível para você.</span><span class="sxs-lookup"><span data-stu-id="5e333-111">Another exciting feature of Script Lab is beta or preview functionality is available for you to try.</span></span>

> [!TIP]
> <span data-ttu-id="5e333-112">Para participar da versão beta ou prévia, talvez seja necessário inscrever-se no [programa Office](https://products.office.com/office-insider)Insider.</span><span class="sxs-lookup"><span data-stu-id="5e333-112">To participate in beta or preview, you may have to sign up for the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="5e333-113">Parece bom até agora?</span><span class="sxs-lookup"><span data-stu-id="5e333-113">Sounds good so far?</span></span> <span data-ttu-id="5e333-114">Dê uma olhada neste vídeo de um minuto para ver o script Lab em ação.</span><span class="sxs-lookup"><span data-stu-id="5e333-114">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="5e333-115">[![Visualizar vídeo mostrando o laboratório de script em execução no Excel, Word e PowerPoint online.] (../images/screenshot-wide-youtube.png 'Vídeo do script Lab Preview')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="5e333-115">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint Online.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="script-lab-supported-clients"></a><span data-ttu-id="5e333-116">Clientes compatíveis com o script Lab</span><span class="sxs-lookup"><span data-stu-id="5e333-116">Script Lab supported clients</span></span>

<span data-ttu-id="5e333-117">O script Lab é compatível com Excel, Word e PowerPoint nos seguintes clientes.</span><span class="sxs-lookup"><span data-stu-id="5e333-117">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="5e333-118">Office no Windows (conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="5e333-118">Office on Windows (connected to Office 365)</span></span>
- <span data-ttu-id="5e333-119">Office para Mac (conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="5e333-119">Office for Mac (connected to Office 365)</span></span>
- <span data-ttu-id="5e333-120">Office Online</span><span class="sxs-lookup"><span data-stu-id="5e333-120">Office Online</span></span>
- <span data-ttu-id="5e333-121">Office 2013 ou posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="5e333-121">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="5e333-122">Office 2016 ou posterior para Mac</span><span class="sxs-lookup"><span data-stu-id="5e333-122">Office 2016 or later for Mac</span></span>

## <a name="next-steps"></a><span data-ttu-id="5e333-123">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="5e333-123">Next steps</span></span>

<span data-ttu-id="5e333-124">Quando estiver pronto para criar seu suplemento do Office, confira o [início rápido de 5 minutos](/office/dev/add-ins/#5-minute-quick-starts) para seu aplicativo preferido do Office.</span><span class="sxs-lookup"><span data-stu-id="5e333-124">When you're ready to create your Office Add-in, see the [5-minute quick start](/office/dev/add-ins/#5-minute-quick-starts) for your preferred Office application.</span></span>

## <a name="see-also"></a><span data-ttu-id="5e333-125">Confira também</span><span class="sxs-lookup"><span data-stu-id="5e333-125">See also</span></span>

- [<span data-ttu-id="5e333-126">Obter o laboratório de scripts</span><span class="sxs-lookup"><span data-stu-id="5e333-126">Get Script Lab</span></span>](https://store.office.com/app.aspx?assetid=WA104380862)
- [<span data-ttu-id="5e333-127">Saiba mais sobre o script Lab</span><span class="sxs-lookup"><span data-stu-id="5e333-127">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="5e333-128">Inscreva-se no programa dev</span><span class="sxs-lookup"><span data-stu-id="5e333-128">Sign up for the dev program</span></span>](https://developer.microsoft.com/office/dev-program)
