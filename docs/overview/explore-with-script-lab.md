---
title: Explorar a API JavaScript do Office usando o script Lab
description: Use o script Lab para explorar a API do Office JS e a funcionalidade de protótipo.
ms.topic: article
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 9ef81443fade450a7bea519a99cb607c320dd4f6
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128640"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="587f9-103">Explorar a API JavaScript do Office usando o script Lab</span><span class="sxs-lookup"><span data-stu-id="587f9-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="587f9-104">O [suplemento de laboratório de script](https://store.office.com/app.aspx?assetid=WA104380862), que está disponível gratuitamente na Office Store, permite explorar a API JavaScript do Office enquanto você estiver trabalhando em um programa do Office, como Excel ou Word.</span><span class="sxs-lookup"><span data-stu-id="587f9-104">The [Script Lab add-in](https://store.office.com/app.aspx?assetid=WA104380862), which is available free from the Office store, enables you to explore the Office JavaScript API while you're working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="587f9-105">O script Lab é uma ferramenta conveniente para adicionar ao seu kit de ferramentas de desenvolvimento conforme você protótipo e verificar a funcionalidade desejada no seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="587f9-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="587f9-106">O que é o script Lab?</span><span class="sxs-lookup"><span data-stu-id="587f9-106">What is Script Lab?</span></span>

<span data-ttu-id="587f9-107">O script Lab é uma ferramenta para qualquer pessoa que deseje saber como desenvolver suplementos do Office usando a API JavaScript do Office no Excel, no Word ou no PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="587f9-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="587f9-108">Ele fornece o IntelliSense para que você possa ver o que está disponível e foi criado na estrutura de Mônaco, a mesma estrutura usada pelo Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="587f9-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="587f9-109">Por meio do laboratório de scripts, você pode acessar uma biblioteca de exemplos para experimentar rapidamente recursos ou pode usar um exemplo como ponto de partida para seu próprio código.</span><span class="sxs-lookup"><span data-stu-id="587f9-109">Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code.</span></span> <span data-ttu-id="587f9-110">Você pode até mesmo usar o script Lab para experimentar as APIs de visualização.</span><span class="sxs-lookup"><span data-stu-id="587f9-110">You can even use Script Lab to try preview APIs.</span></span>

<span data-ttu-id="587f9-111">Parece bom até agora?</span><span class="sxs-lookup"><span data-stu-id="587f9-111">Sounds good so far?</span></span> <span data-ttu-id="587f9-112">Dê uma olhada neste vídeo de um minuto para ver o script Lab em ação.</span><span class="sxs-lookup"><span data-stu-id="587f9-112">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="587f9-113">[![Visualizar vídeo mostrando o laboratório de script em execução no Excel, Word e PowerPoint.] (../images/screenshot-wide-youtube.png 'Vídeo do script Lab Preview')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="587f9-113">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="key-features"></a><span data-ttu-id="587f9-114">Principais recursos</span><span class="sxs-lookup"><span data-stu-id="587f9-114">Key features</span></span>

<span data-ttu-id="587f9-115">O script Lab oferece vários recursos para ajudá-lo a explorar a API JavaScript do Office e a funcionalidade do suplemento de protótipo.</span><span class="sxs-lookup"><span data-stu-id="587f9-115">Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.</span></span>

### <a name="explore-samples"></a><span data-ttu-id="587f9-116">Explorar exemplos</span><span class="sxs-lookup"><span data-stu-id="587f9-116">Explore samples</span></span>

<span data-ttu-id="587f9-117">Comece rapidamente com uma coleção de trechos de código internos que mostram como concluir determinadas tarefas com a API.</span><span class="sxs-lookup"><span data-stu-id="587f9-117">Get started quickly with a collection of built-in sample snippets that show how to complete certain tasks with the API.</span></span> <span data-ttu-id="587f9-118">Você pode executar os exemplos para ver instantaneamente o resultado no painel de tarefas ou no documento, examinar os exemplos para saber como a API funciona e até mesmo usar trechos de exemplo como base para a funcionalidade de protótipo do seu próprio suplemento.</span><span class="sxs-lookup"><span data-stu-id="587f9-118">You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use sample snippets as the basis for prototyping functionality of your own add-in.</span></span>

![Exemplos](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a><span data-ttu-id="587f9-120">Código e estilo</span><span class="sxs-lookup"><span data-stu-id="587f9-120">Code and style</span></span>

<span data-ttu-id="587f9-121">Além do código JavaScript ou TypeScript que chama a API do Office JS, cada trecho também contém marcação HTML que define o conteúdo do painel de tarefas e o CSS que define a aparência do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="587f9-121">In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane.</span></span> <span data-ttu-id="587f9-122">Você pode personalizar a marcação HTML e o CSS para testar o posicionamento e o estilo do elemento conforme o design do painel de tarefas do protótipo para seu próprio suplemento.</span><span class="sxs-lookup"><span data-stu-id="587f9-122">You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.</span></span>

> [!TIP]
> <span data-ttu-id="587f9-123">Para chamar APIs de visualização dentro de um trecho de código, você precisará atualizar as bibliotecas do trecho de código para`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`usar a CDN beta () `@types/office-js-preview`e as definições de tipo de visualização.</span><span class="sxs-lookup"><span data-stu-id="587f9-123">To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`.</span></span> <span data-ttu-id="587f9-124">Além disso, algumas APIs de visualização são acessíveis somente se você se inscreveu no [programa Office](https://products.office.com/office-insider) Insider e está executando uma compilação do Office Insider.</span><span class="sxs-lookup"><span data-stu-id="587f9-124">Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://products.office.com/office-insider) and are running an Insider build of Office.</span></span>

### <a name="save-and-share-snippets"></a><span data-ttu-id="587f9-125">Salvar e compartilhar trechos de código</span><span class="sxs-lookup"><span data-stu-id="587f9-125">Save and share snippets</span></span>

<span data-ttu-id="587f9-126">Por padrão, os trechos de código abertos no laboratório de script serão salvos no cache do navegador.</span><span class="sxs-lookup"><span data-stu-id="587f9-126">By default, snippets that you open in Script Lab will be saved to your browser cache.</span></span> <span data-ttu-id="587f9-127">Para salvar um trecho permanentemente, você pode exportá-lo para um [GitHub](https://gist.github.com).</span><span class="sxs-lookup"><span data-stu-id="587f9-127">To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com).</span></span> <span data-ttu-id="587f9-128">Crie uma propriedade secreta para salvar um trecho de código exclusivamente para uso próprio ou criar uma propriedade compartilhada se você planeja compartilhá-la com outras pessoas.</span><span class="sxs-lookup"><span data-stu-id="587f9-128">Create a secret gist to save a snippet exclusively for your own use, or create a shared gist if you plan to share it with others.</span></span>

![Opções de compartilhamento](../images/script-lab-share.jpg)

### <a name="import-snippets"></a><span data-ttu-id="587f9-130">Importar trechos</span><span class="sxs-lookup"><span data-stu-id="587f9-130">Import snippets</span></span>

<span data-ttu-id="587f9-131">Você pode importar um trecho para o laboratório de script especificando a URL para o membro do [GitHub](https://gist.github.com) público onde o YAML de trecho de código está armazenado ou colando no YAML completo para o trecho de código.</span><span class="sxs-lookup"><span data-stu-id="587f9-131">You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet.</span></span> <span data-ttu-id="587f9-132">Esse recurso pode ser útil em situações em que alguém compartilhou seus trechos de código com você publicando-o em um próprio GitHub ou fornecendo a YAML de seus trechos de código.</span><span class="sxs-lookup"><span data-stu-id="587f9-132">This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.</span></span>

![Opção importar trecho](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a><span data-ttu-id="587f9-134">Clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="587f9-134">Supported clients</span></span>

<span data-ttu-id="587f9-135">O script Lab é compatível com Excel, Word e PowerPoint nos seguintes clientes.</span><span class="sxs-lookup"><span data-stu-id="587f9-135">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="587f9-136">Office 2013 ou posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="587f9-136">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="587f9-137">Office 2016 ou posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="587f9-137">Office 2016 or later on Mac</span></span>
- <span data-ttu-id="587f9-138">Office na Web</span><span class="sxs-lookup"><span data-stu-id="587f9-138">Office on the web</span></span>

## <a name="next-steps"></a><span data-ttu-id="587f9-139">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="587f9-139">Next steps</span></span>

<span data-ttu-id="587f9-140">Você é bem-vindo à expansão da biblioteca de exemplo no laboratório de scripts, contribuindo novos trechos de código para o repositório do GitHub [Office-js-Snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) .</span><span class="sxs-lookup"><span data-stu-id="587f9-140">You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.</span></span>

<span data-ttu-id="587f9-141">Quando estiver pronto para criar seu suplemento do Office, confira o [início rápido de 5 minutos](/office/dev/add-ins/#5-minute-quick-starts) para seu aplicativo preferido do Office.</span><span class="sxs-lookup"><span data-stu-id="587f9-141">When you're ready to create your Office Add-in, see the [5-minute quick start](/office/dev/add-ins/#5-minute-quick-starts) for your preferred Office application.</span></span>

## <a name="see-also"></a><span data-ttu-id="587f9-142">Confira também</span><span class="sxs-lookup"><span data-stu-id="587f9-142">See also</span></span>

- [<span data-ttu-id="587f9-143">Obter o laboratório de scripts</span><span class="sxs-lookup"><span data-stu-id="587f9-143">Get Script Lab</span></span>](https://store.office.com/app.aspx?assetid=WA104380862)
- [<span data-ttu-id="587f9-144">Saiba mais sobre o script Lab</span><span class="sxs-lookup"><span data-stu-id="587f9-144">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="587f9-145">Inscreva-se no programa dev</span><span class="sxs-lookup"><span data-stu-id="587f9-145">Sign up for the dev program</span></span>](https://developer.microsoft.com/office/dev-program)
