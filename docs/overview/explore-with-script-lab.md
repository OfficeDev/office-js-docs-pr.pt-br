---
title: Explore a API JavaScript do Office usando o Script Lab
description: Use o script Lab para explorar a funcionalidade de protótipo e a API do Office JS.
ms.date: 04/16/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 6fb886f1c86267ed7081d1892d1314798ab4cedc
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547252"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="29285-103">Explore a API JavaScript do Office usando o Script Lab</span><span class="sxs-lookup"><span data-stu-id="29285-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="29285-104">O[Script Lab é um suplemento](https://appsource.microsoft.com/product/office/WA104380862), que está disponível gratuitamente em AppSource, permite que você explore a API JavaScript do Office enquanto você trabalha em um programa do Office, como o Excel ou o Word.</span><span class="sxs-lookup"><span data-stu-id="29285-104">The [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862), which is available free from AppSource, enables you to explore the Office JavaScript API while you're working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="29285-105">O script Lab é uma ferramenta conveniente para adicionar ao seu kit de ferramentas de desenvolvimento durante o protótipo e a verificação da funcionalidade que você deseja em seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="29285-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="29285-106">O que é o script Lab?</span><span class="sxs-lookup"><span data-stu-id="29285-106">What is Script Lab?</span></span>

<span data-ttu-id="29285-107">O Script Lab é uma ferramenta para qualquer pessoa que queira aprender a desenvolver suplementos do Office usando a API do JavaScript do Office no Excel, Word ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="29285-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="29285-108">Ele fornece IntelliSense para que você possa ver o que está disponível e que foi criado na estrutura de Mônaco, a mesma estrutura usada pelo código do Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="29285-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="29285-109">Por meio do Script Lab, você pode acessar uma biblioteca de amostras para experimentar rapidamente recursos ou até mesmo usar um exemplo como o ponto de partida para o seu próprio código.</span><span class="sxs-lookup"><span data-stu-id="29285-109">Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code.</span></span> <span data-ttu-id="29285-110">Você pode até usar o Script Lab para experimentar as APIs de visualização.</span><span class="sxs-lookup"><span data-stu-id="29285-110">You can even use Script Lab to try preview APIs.</span></span>

<span data-ttu-id="29285-111">Parece bom?</span><span class="sxs-lookup"><span data-stu-id="29285-111">Sounds good so far?</span></span> <span data-ttu-id="29285-112">Dê uma olhada neste vídeo de um minuto para ver Script Lab em ação.</span><span class="sxs-lookup"><span data-stu-id="29285-112">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="29285-113">[![Visualização de vídeo mostrando o Script Lab em execução no Excel, Word e PowerPoint.](../images/screenshot-wide-youtube.png 'Visualização de vídeo do Script Lab')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="29285-113">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="key-features"></a><span data-ttu-id="29285-114">Principais recursos</span><span class="sxs-lookup"><span data-stu-id="29285-114">Key features</span></span>

<span data-ttu-id="29285-115">O script Lab oferece vários recursos para ajudá-lo a explorar a funcionalidade do suplemento API e protótipo do Office JavaScript.</span><span class="sxs-lookup"><span data-stu-id="29285-115">Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.</span></span>

### <a name="explore-samples"></a><span data-ttu-id="29285-116">Explorar amostras</span><span class="sxs-lookup"><span data-stu-id="29285-116">Explore samples</span></span>

<span data-ttu-id="29285-117">Comece a trabalhar rapidamente com um conjunto de exemplos internos que mostram como concluir tarefas com a API.</span><span class="sxs-lookup"><span data-stu-id="29285-117">Get started quickly with a collection of built-in sample snippets that show how to complete tasks with the API.</span></span> <span data-ttu-id="29285-118">Você pode executar as amostras para ver instantaneamente o resultado no painel de tarefas ou documento, examinar os exemplos para saber como a API funciona e até mesmo usar amostras para criar um protótipo do seu próprio suplemento.</span><span class="sxs-lookup"><span data-stu-id="29285-118">You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

![Exemplos](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a><span data-ttu-id="29285-120">Código e estilo</span><span class="sxs-lookup"><span data-stu-id="29285-120">Code and style</span></span>

<span data-ttu-id="29285-121">Além de código JavaScript ou TypeScript que chama a API do Office JS, cada snippet também contém marcação HTML que define o conteúdo do painel de tarefas e CSS que define a aparência do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="29285-121">In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane.</span></span> <span data-ttu-id="29285-122">Você pode personalizar a marcação HTML e CSS para experimentar o posicionamento e o estilo de elementos à medida que você cria seu próprio suplemento no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="29285-122">You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.</span></span>

> [!TIP]
> <span data-ttu-id="29285-123">Para chamar as APIs de visualização dentro de um snippet, você precisará atualizar as bibliotecas do trecho para usar a CDN beta (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) e as `@types/office-js-preview`definições de tipo de visualização.</span><span class="sxs-lookup"><span data-stu-id="29285-123">To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`.</span></span> <span data-ttu-id="29285-124">Além disso, algumas APIs de visualização são acessíveis apenas se você se inscreveu no programa [Office Insider](https://insider.office.com) e está executando uma compilação do Office Insider.</span><span class="sxs-lookup"><span data-stu-id="29285-124">Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://insider.office.com) and are running an Insider build of Office.</span></span>

### <a name="save-and-share-snippets"></a><span data-ttu-id="29285-125">Salvar e compartilhar trechos</span><span class="sxs-lookup"><span data-stu-id="29285-125">Save and share snippets</span></span>

<span data-ttu-id="29285-126">Por padrão, os trechos abertos no Script Lab serão salvos no cache do navegador.</span><span class="sxs-lookup"><span data-stu-id="29285-126">By default, snippets that you open in Script Lab will be saved to your browser cache.</span></span> <span data-ttu-id="29285-127">Para salvar um trecho permanentemente, você pode exportá-lo para um [GitHub gist](https://gist.github.com).</span><span class="sxs-lookup"><span data-stu-id="29285-127">To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com).</span></span> <span data-ttu-id="29285-128">Crie uma propriedade secreta para salvar um trecho exclusivo para seu próprio uso ou criar uma conta pública se planejar compartilhá-la com outras pessoas.</span><span class="sxs-lookup"><span data-stu-id="29285-128">Create a secret gist to save a snippet exclusively for your own use, or create a public gist if you plan to share it with others.</span></span>

![Opções de compartilhamento](../images/script-lab-share.jpg)

### <a name="import-snippets"></a><span data-ttu-id="29285-130">Importar trechos</span><span class="sxs-lookup"><span data-stu-id="29285-130">Import snippets</span></span>

<span data-ttu-id="29285-131">Você pode importar um trecho para o Script Lab especificando a URL para o [do GitHub público](https://gist.github.com) onde o snippet YAML está armazenado ou colando-o no YAML completo do trecho.</span><span class="sxs-lookup"><span data-stu-id="29285-131">You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet.</span></span> <span data-ttu-id="29285-132">Esse recurso pode ser útil em situações em que outra pessoa compartilhou trechos com você publicando-o em uma oferta do GitHub ou fornecendo o YAML do trecho.</span><span class="sxs-lookup"><span data-stu-id="29285-132">This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.</span></span>

![Opção importar trecho](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a><span data-ttu-id="29285-134">Clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="29285-134">Supported clients</span></span>

<span data-ttu-id="29285-135">O Script Lab tem suporte para o Excel, o Word e o PowerPoint nos clientes a seguir.</span><span class="sxs-lookup"><span data-stu-id="29285-135">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="29285-136">Office 2013 ou posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="29285-136">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="29285-137">Office 2016 ou posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="29285-137">Office 2016 or later on Mac</span></span>
- <span data-ttu-id="29285-138">Office na Web</span><span class="sxs-lookup"><span data-stu-id="29285-138">Office on the web</span></span>

## <a name="next-steps"></a><span data-ttu-id="29285-139">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="29285-139">Next steps</span></span>

<span data-ttu-id="29285-140">Para usar o Script Lab no Excel, no Word ou no PowerPoint, instale o [suplemento do Script Lab](https://appsource.microsoft.com/product/office/WA104380862) do AppSource.</span><span class="sxs-lookup"><span data-stu-id="29285-140">To use Script Lab in Excel, Word, or PowerPoint, install the [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862) from AppSource.</span></span> 

<span data-ttu-id="29285-141">Você é bem-vindo a expandir a biblioteca de exemplo no Script Lab, contribuindo com novos trechos para o [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) repositório do GitHub.</span><span class="sxs-lookup"><span data-stu-id="29285-141">You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.</span></span>

<span data-ttu-id="29285-142">Quando estiver pronto para criar seu primeiro suplemento do Office, experimente o início rápido para [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](../quickstarts/outlook-quickstart.md), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md)ou [Project](../quickstarts/project-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="29285-142">When you're ready to create your first Office Add-in, try out the quick start for [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](../quickstarts/outlook-quickstart.md), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md), or [Project](../quickstarts/project-quickstart.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="29285-143">Confira também</span><span class="sxs-lookup"><span data-stu-id="29285-143">See also</span></span>

- [<span data-ttu-id="29285-144">Obter Script Lab</span><span class="sxs-lookup"><span data-stu-id="29285-144">Get Script Lab</span></span>](https://appsource.microsoft.com/product/office/WA104380862)
- [<span data-ttu-id="29285-145">Saiba mais sobre o Script Lab</span><span class="sxs-lookup"><span data-stu-id="29285-145">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="29285-146">Ingressar no Programa para Desenvolvedores do Office 365</span><span class="sxs-lookup"><span data-stu-id="29285-146">Join the Office 365 Developer Program</span></span>](https://developer.microsoft.com/office/dev-program)
- [<span data-ttu-id="29285-147">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="29285-147">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
