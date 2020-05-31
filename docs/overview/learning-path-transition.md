---
title: Transição aqui! Um guia para criadores de suplemento do VSTO que fazem suplementos Web do Office
description: Um roteiro recomendado para desenvolvedores experientes de suplemento do VSTO para recursos de aprendizagem de suplementos Web do Office.
ms.date: 05/10/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 6ed812bae73282999716c448ef683dcc6aeae778
ms.sourcegitcommit: 735bf94ac3c838f580a992e7ef074dbc8be2b0ea
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/08/2020
ms.locfileid: "44170832"
---
# <a name="transition-here-a-guide-for-vsto-add-in-creators-making-office-web-add-ins"></a><span data-ttu-id="5abc0-104">Transição aqui!</span><span class="sxs-lookup"><span data-stu-id="5abc0-104">Transition Here!</span></span> <span data-ttu-id="5abc0-105">Um guia para criadores de suplemento do VSTO que fazem suplementos Web do Office</span><span class="sxs-lookup"><span data-stu-id="5abc0-105">A guide for VSTO add-in creators making Office Web Add-ins</span></span>

<span data-ttu-id="5abc0-106">Você criou alguns suplementos do VSTO para aplicativos do Office executados no Windows, e agora está aprendendo um nova maneira de estender o Office que será executado no Windows, no Mac e na versão online do pacote do Office: suplementos Web do Office.</span><span class="sxs-lookup"><span data-stu-id="5abc0-106">So, you've made some VSTO add-ins for Office applications that run on Windows and now you're exploring the new way of extending Office that will run on Windows, Mac, and the online version of the Office suite: Office Web Add-ins.</span></span>

<span data-ttu-id="5abc0-107">Sua compreensão sobre os modelos de objeto para Excel, Word e outros aplicativos do Office será uma grande ajuda, pois os modelos de objeto nos suplementos Web do Office seguem padrões semelhantes.</span><span class="sxs-lookup"><span data-stu-id="5abc0-107">Your understanding of the object models for the Excel, Word, and the other Office applications will be a huge help because the object models in Office Web Add-ins follow similar patterns.</span></span> <span data-ttu-id="5abc0-108">Mas haverá alguns desafios:</span><span class="sxs-lookup"><span data-stu-id="5abc0-108">But there are going to be some challenges:</span></span>

- <span data-ttu-id="5abc0-109">Você trabalhará com uma linguagem diferente (JavaScript ou TypeScript) em vez de C# ou Visual Basic .NET.</span><span class="sxs-lookup"><span data-stu-id="5abc0-109">You will be working with a different language (either JavaScript or TypeScript) instead of C# or Visual Basic .NET.</span></span> <span data-ttu-id="5abc0-110">(Há também uma maneira, descrita abaixo, de reutilizar alguns de seus códigos existentes em um suplemento Web).</span><span class="sxs-lookup"><span data-stu-id="5abc0-110">(There is also a way, described below, to reuse some of your existing code in a web add-in.)</span></span>
- <span data-ttu-id="5abc0-111">Os suplementos Web do Office são implantados de forma diferente dos suplementos do VSTO.</span><span class="sxs-lookup"><span data-stu-id="5abc0-111">Office Web Add-ins are deployed differently from VSTO add-ins.</span></span>
- <span data-ttu-id="5abc0-112">Os suplementos Web do Office são aplicativos Web executados em uma janela simplificada do navegador que está incorporada ao aplicativo do Office. Portanto, é necessário obter um conhecimento básico dos aplicativos Web e de como eles são hospedados em servidores Web ou em contas de nuvem.</span><span class="sxs-lookup"><span data-stu-id="5abc0-112">Office Web Add-ins are web applications that run in a simplified browser window that is embedded in the Office application, so you need to gain a basic understanding of web applications and how they are hosted on web servers or cloud accounts.</span></span> 

<span data-ttu-id="5abc0-113">Por esses motivos, grande parte deste artigo duplica nosso roteiro de aprendizagem para iniciantes das extensões do Office: [Comece aqui! Um guia para iniciantes na criação de Suplementos do Office](learning-path-beginner.md). O que adicionamos são alguns recursos extras de aprendizagem para ajudar os desenvolvedores de suplemento do VSTO a aproveitar suas experiências, e também a reutilizar códigos existentes.</span><span class="sxs-lookup"><span data-stu-id="5abc0-113">For these reasons, much of this article duplicates our learning path for complete beginners to Office extensions: [Start Here! A guide for beginners making Office Add-ins](learning-path-beginner.md). What we have added are some additional learning resources to help VSTO add-in developers leverage their experience, and also help them reuse their existing code.</span></span>

## <a name="step-0-prerequisites"></a><span data-ttu-id="5abc0-114">Etapa 0: Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="5abc0-114">Step 0: Prerequisites</span></span>

- <span data-ttu-id="5abc0-115">Os suplementos Web do Office (também chamados de suplementos do Office) são essencialmente aplicativos Web incorporados no Office.</span><span class="sxs-lookup"><span data-stu-id="5abc0-115">Office Web Add-ins (also referred to as Office Add-ins) are essentially web applications embedded in Office.</span></span> <span data-ttu-id="5abc0-116">Portanto, você deve primeiro ter um conhecimento básico dos aplicativos Web e de como eles são hospedados na Web.</span><span class="sxs-lookup"><span data-stu-id="5abc0-116">So, you should first have a basic understanding of web applications and how they are hosted on the web.</span></span> <span data-ttu-id="5abc0-117">Há uma quantidade enorme de informações sobre isso na Internet, em livros e em cursos online.</span><span class="sxs-lookup"><span data-stu-id="5abc0-117">There's an enormous amount of information about this on the Internet, in books, and in online courses.</span></span> <span data-ttu-id="5abc0-118">Uma boa maneira de começar, se você não tem nenhum conhecimento prévio sobre aplicativos da Web, é procurar "O que é um aplicativo da Web?"</span><span class="sxs-lookup"><span data-stu-id="5abc0-118">A good way to start if you have no prior knowledge of web applications at all is to search for "What is a web app?"</span></span> <span data-ttu-id="5abc0-119">no Bing.</span><span class="sxs-lookup"><span data-stu-id="5abc0-119">on Bing.</span></span>
- <span data-ttu-id="5abc0-120">A principal linguagem de programação que você usará na criação de suplementos do Office é o JavaScript ou o TypeScript.</span><span class="sxs-lookup"><span data-stu-id="5abc0-120">The primary programming language you'll use to create Office Add-ins is JavaScript or TypeScript.</span></span> <span data-ttu-id="5abc0-121">Pense no TypeScript como uma versão fortemente tipada do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="5abc0-121">You can think of TypeScript as a strongly-typed version of JavaScript.</span></span> <span data-ttu-id="5abc0-122">Se você não conhece nenhuma dessas linguagens, mas tem experiência com VBA, VB.Net e C#, provavelmente achará o TypeScript mais fácil de aprender.</span><span class="sxs-lookup"><span data-stu-id="5abc0-122">If you're not familiar with either of these languages, but you have experience with VBA, VB.Net, C#, you'll probably find TypeScript easier to learn.</span></span> <span data-ttu-id="5abc0-123">Novamente, há muitas informações sobre essas linguagens de programação na Internet, em livros e em cursos online.</span><span class="sxs-lookup"><span data-stu-id="5abc0-123">Again, there's a wealth of information about these languages on the Internet, in books, and in online courses.</span></span>

## <a name="step-1-begin-with-fundamentals"></a><span data-ttu-id="5abc0-124">Etapa 1: Comece com os fundamentos</span><span class="sxs-lookup"><span data-stu-id="5abc0-124">Step 1: Begin with fundamentals</span></span>

<span data-ttu-id="5abc0-125">Sabemos que você está ansioso para começar a codificar, mas há algumas coisas sobre os Suplementos do Office que você deve ler antes de abrir o IDE ou o editor de código.</span><span class="sxs-lookup"><span data-stu-id="5abc0-125">We know you're eager to start coding, but there are some things about Office Add-ins that you should read before you open your IDE or code editor.</span></span>

- <span data-ttu-id="5abc0-126">[Visão Geral da Plataforma de Suplementos do Office](office-add-ins.md): Descubra o que são os suplementos da Web do Office e como eles diferem das formas mais antigas de estender o Office, como os suplementos do VSTO.</span><span class="sxs-lookup"><span data-stu-id="5abc0-126">[Office Add-ins Platform Overview](office-add-ins.md): Find out what Office Web Add-ins are and how they differ from older ways of extending Office, such as VSTO add-ins.</span></span>
- <span data-ttu-id="5abc0-127">[Criação de Suplementos do Office](office-add-ins-fundamentals.md): Obtenha uma visão geral do desenvolvimento e do ciclo de vida de suplementos do Office, incluindo ferramentas, criação de uma Interface de Usuário do suplemento e uso das APIs JavaScript para interagir com o documento do Office.</span><span class="sxs-lookup"><span data-stu-id="5abc0-127">[Building Office Add-ins](office-add-ins-fundamentals.md): Get an overview of Office add-in development and lifecycle including tooling, creating an add-in UI, and using the JavaScript APIs to interact with the Office document.</span></span>

<span data-ttu-id="5abc0-128">Existem muitos links nesses artigos, mas se você estiver migrando para os suplementos Web do Office, recomendamos que você volte aqui quando os tiver lido e continue na próxima seção.</span><span class="sxs-lookup"><span data-stu-id="5abc0-128">There are a lot of links in those articles, but if you're transitioning to Office Web Add-ins, we recommend that you come back here when you've read them and continue with the next section.</span></span>

## <a name="step-2-install-tools-and-create-your-first-add-in"></a><span data-ttu-id="5abc0-129">Etapa 2: Instale ferramentas e crie o seu primeiro suplemento</span><span class="sxs-lookup"><span data-stu-id="5abc0-129">Step 2: Install tools and create your first add-in</span></span>

<span data-ttu-id="5abc0-130">Agora você tem uma visão geral, então comece com um de nossos inícios rápidos.</span><span class="sxs-lookup"><span data-stu-id="5abc0-130">You've got the big picture now, so dive in with one of our quick starts.</span></span> <span data-ttu-id="5abc0-131">Para fins de aprendizado da plataforma, recomendamos o início rápido do Excel.</span><span class="sxs-lookup"><span data-stu-id="5abc0-131">For purposes of learning the platform, we recommend the Excel quick start.</span></span> <span data-ttu-id="5abc0-132">Há uma versão baseada no Visual Studio e outra baseada em Node.js e Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="5abc0-132">There's a version based on Visual Studio and another based on Node.js and Visual Studio Code.</span></span> <span data-ttu-id="5abc0-133">Se você estiver migrando de suplementos do VSTO, provavelmente encontrará a versão do Visual Studio mais fácil de trabalhar.</span><span class="sxs-lookup"><span data-stu-id="5abc0-133">If you're transitioning from VSTO add-ins, you'll probably find the Visual Studio version easier to work with.</span></span>

- [<span data-ttu-id="5abc0-134">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="5abc0-134">Visual Studio</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [<span data-ttu-id="5abc0-135">Node.js e Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="5abc0-135">Node.js and Visual Studio Code</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a><span data-ttu-id="5abc0-136">Etapa 3: Codifique</span><span class="sxs-lookup"><span data-stu-id="5abc0-136">Step 3: Code</span></span>

<span data-ttu-id="5abc0-137">Não se pode aprender a dirigir lendo o manual do proprietário, então comece a codificar com este [tutorial do Excel](../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="5abc0-137">You can't learn to drive by reading the owner's manual, so start coding with this [Excel tutorial](../tutorials/excel-tutorial.md).</span></span> <span data-ttu-id="5abc0-138">Você usará a biblioteca JavaScript do Office e um pouco de XML no manifesto dos suplementos.</span><span class="sxs-lookup"><span data-stu-id="5abc0-138">You'll be using the Office JavaScript library and some XML in the add-in's manifest.</span></span> <span data-ttu-id="5abc0-139">Não é necessário memorizar nada, porque você terá mais informações sobre ambos em etapas posteriores.</span><span class="sxs-lookup"><span data-stu-id="5abc0-139">There's no need to memorize anything, because you'll be getting more background about both in a later step.</span></span>

## <a name="step-4-understand-the-javascript-library"></a><span data-ttu-id="5abc0-140">Etapa 4: Entenda a biblioteca JavaScript</span><span class="sxs-lookup"><span data-stu-id="5abc0-140">Step 4: Understand the JavaScript library</span></span>

<span data-ttu-id="5abc0-141">Obtenha uma visão geral da biblioteca JavaScript do Office com este tutorial do Microsoft Learn: [Entenda as APIs JavaScript do Office](/learn/modules/intro-office-add-ins/3-apis).</span><span class="sxs-lookup"><span data-stu-id="5abc0-141">Get the big picture of the Office JavaScript library with this tutorial from Microsoft Learn: [Understand the Office JavaScript APIs](/learn/modules/intro-office-add-ins/3-apis).</span></span>

<span data-ttu-id="5abc0-142">Em seguida, explore as APIs do Office JavaScript com a [ferramenta Script Lab](explore-with-script-lab.md) – uma área restrita para executar e explorar as APIs.</span><span class="sxs-lookup"><span data-stu-id="5abc0-142">Then explore the Office JavaScript APIs with the [Script Lab tool](explore-with-script-lab.md) -- a sandbox for running and exploring the APIs.</span></span>

### <a name="special-resource-for-vsto-add-in-developers"></a><span data-ttu-id="5abc0-143">Um recurso especial para desenvolvedores de suplemento do VSTO</span><span class="sxs-lookup"><span data-stu-id="5abc0-143">Special resource for VSTO add-in developers</span></span>

<span data-ttu-id="5abc0-144">Esse seria um bom lugar para dar uma olhada no exemplo de suplemento, [Suplemento do Excel JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).</span><span class="sxs-lookup"><span data-stu-id="5abc0-144">This would be a good place to take a look at the sample add-in, [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).</span></span> <span data-ttu-id="5abc0-145">Ele foi criado para destacar as semelhanças e diferenças entre suplementos do VSTO e suplementos Web do Office, e o leiame do exemplo indica os pontos importantes da comparação.</span><span class="sxs-lookup"><span data-stu-id="5abc0-145">It was created to highlight the similarities and differences between VSTO add-ins and Office Web Add-ins, and the readme of the sample calls out the important points of comparison.</span></span>

## <a name="step-5-understand-the-manifest"></a><span data-ttu-id="5abc0-146">Etapa 5: Entenda o manifesto</span><span class="sxs-lookup"><span data-stu-id="5abc0-146">Step 5: Understand the manifest</span></span>

<span data-ttu-id="5abc0-147">Entenda os objetivos do manifesto de suplemento Web e veja uma introdução à sua marcação XML no [Manifesto XML dos suplementos do Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="5abc0-147">Get an understanding of the purposes of the web add-in manifest and an introduction to its XML markup in [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

## <a name="step-6-for-vsto-developers-only-reuse-your-vsto-code"></a><span data-ttu-id="5abc0-148">Etapa 6 (somente para desenvolvedores do VSTO): Reutilize seu código de VSTO</span><span class="sxs-lookup"><span data-stu-id="5abc0-148">Step 6 (for VSTO developers only): Reuse your VSTO code</span></span>

<span data-ttu-id="5abc0-149">Você pode reutilizar alguns dos códigos de suplemento do VSTO em um suplemento Web do Office, movendo-os para o back-end do seu aplicativo Web no servidor e disponibilizando-o para o JavaScript ou TypeScript como uma API da Web.</span><span class="sxs-lookup"><span data-stu-id="5abc0-149">You can reuse some of your VSTO add-in code in an Office web add-in by moving it to your web application's back end on the server and making it available to your JavaScript or TypeScript as a web API.</span></span> <span data-ttu-id="5abc0-150">Para obter instruções, confira [Tutorial: compartilhar código entre um suplemento do VSTO e um suplemento do Office usando uma biblioteca de códigos compartilhados](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="5abc0-150">For guidance, see [Tutorial: Share code between both a VSTO Add-in and an Office add-in by using a shared code library](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="5abc0-151">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="5abc0-151">Next Steps</span></span>

<span data-ttu-id="5abc0-152">Parabéns por concluir o roteiro de aprendizagem para desenvolvedores de suplementos VSTO para suplementos Web do Office!</span><span class="sxs-lookup"><span data-stu-id="5abc0-152">Congratulations on finishing the VSTO add-in developer's learning path for Office Web Add-ins!</span></span> <span data-ttu-id="5abc0-153">Veja algumas sugestões para explorar ainda mais a documentação:</span><span class="sxs-lookup"><span data-stu-id="5abc0-153">Here are some suggestions for further exploration of our documentation:</span></span>

- <span data-ttu-id="5abc0-154">Tutoriais ou inícios rápidos para outros aplicativos do Office:</span><span class="sxs-lookup"><span data-stu-id="5abc0-154">Tutorials or quick starts for other Office applications:</span></span>

  - [<span data-ttu-id="5abc0-155">Início rápido do OneNote</span><span class="sxs-lookup"><span data-stu-id="5abc0-155">OneNote quick start</span></span>](../quickstarts/onenote-quickstart.md)
  - [<span data-ttu-id="5abc0-156">Tutorial do Outlook</span><span class="sxs-lookup"><span data-stu-id="5abc0-156">Outlook tutorial</span></span>](/outlook/add-ins/addin-tutorial)
  - [<span data-ttu-id="5abc0-157">Tutorial do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5abc0-157">PowerPoint tutorial</span></span>](../tutorials/powerpoint-tutorial.md)
  - [<span data-ttu-id="5abc0-158">Início rápido do Project</span><span class="sxs-lookup"><span data-stu-id="5abc0-158">Project quick start</span></span>](../quickstarts/project-quickstart.md)
  - [<span data-ttu-id="5abc0-159">Tutorial do Word</span><span class="sxs-lookup"><span data-stu-id="5abc0-159">Word tutorial</span></span>](../tutorials/word-tutorial.md)

- <span data-ttu-id="5abc0-160">Outros assuntos importantes:</span><span class="sxs-lookup"><span data-stu-id="5abc0-160">Other important subjects:</span></span>

  - [<span data-ttu-id="5abc0-161">Desenvolver Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="5abc0-161">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
  - [<span data-ttu-id="5abc0-162">Práticas recomendadas para o desenvolvimento de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5abc0-162">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
  - [<span data-ttu-id="5abc0-163">Fazer o design de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5abc0-163">Design Office Add-ins</span></span>](../design/add-in-design.md)
  - [<span data-ttu-id="5abc0-164">Testar e depurar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5abc0-164">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
  - [<span data-ttu-id="5abc0-165">Implantar e publicar Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5abc0-165">Deploy and publish Office Add-ins</span></span>](../publish/publish.md)
  - [<span data-ttu-id="5abc0-166">Recursos</span><span class="sxs-lookup"><span data-stu-id="5abc0-166">Resources</span></span>](../resources/resources-links-help.md)