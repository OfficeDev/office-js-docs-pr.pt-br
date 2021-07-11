---
title: Carregar o ambiente de tempo de execução e DOM
description: Carregue o dom e Office ambiente de tempo de execução de complementos.
ms.date: 04/20/2021
localization_priority: Normal
ms.openlocfilehash: 0cfdcf3750d9c0a3dd21667729da59dbfedf61c8
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349837"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="0a411-103">Carregar o ambiente de tempo de execução e DOM</span><span class="sxs-lookup"><span data-stu-id="0a411-103">Loading the DOM and runtime environment</span></span>

<span data-ttu-id="0a411-104">Um suplemento deve garantir que o DOM e o ambiente de tempo de execução de Suplementos do Office sejam carregados antes de executar sua própria lógica personalizada.</span><span class="sxs-lookup"><span data-stu-id="0a411-104">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span>

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="0a411-105">Inicialização de um suplemento de conteúdo ou de painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="0a411-105">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="0a411-106">A figura a seguir mostra o fluxo de eventos envolvidos na inicialização de um suplemento de conteúdo ou de painel de tarefas no Excel, no PowerPoint, no Project ou no Word.</span><span class="sxs-lookup"><span data-stu-id="0a411-106">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, or Word.</span></span>

![Flow de eventos ao iniciar um conteúdo ou um complemento do painel de tarefas.](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="0a411-108">Os eventos a seguir ocorrem quando um conteúdo ou um complemento do painel de tarefas é iniciado.</span><span class="sxs-lookup"><span data-stu-id="0a411-108">The following events occur when a content or task pane add-in starts.</span></span>

1. <span data-ttu-id="0a411-109">O usuário abre um documento que já contém um suplemento ou insere um suplemento no documento.</span><span class="sxs-lookup"><span data-stu-id="0a411-109">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>

2. <span data-ttu-id="0a411-110">O Office de cliente lê o manifesto XML do add-in do AppSource, um catálogo de aplicativos no SharePoint ou o catálogo de pastas compartilhadas de onde ele se origina.</span><span class="sxs-lookup"><span data-stu-id="0a411-110">The Office client application reads the add-in's XML manifest from AppSource, an app catalog on SharePoint, or the shared folder catalog it originates from.</span></span>

3. <span data-ttu-id="0a411-111">O Office cliente abre a página HTML do complemento em um controle do navegador.</span><span class="sxs-lookup"><span data-stu-id="0a411-111">The Office client application opens the add-in's HTML page in a browser control.</span></span>

    <span data-ttu-id="0a411-p101">As próximas duas etapas, as etapas 4 e 5, ocorrem de forma assíncrona e em paralelo. Por esse motivo, o código do suplemento deve garantir que o DOM e o ambiente do tempo de execução do suplemento tenham terminado de carregar antes de prosseguir.</span><span class="sxs-lookup"><span data-stu-id="0a411-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>

4. <span data-ttu-id="0a411-114">O controle do navegador carrega o corpo DOM e HTML e chama o manipulador de eventos para o `window.onload` evento.</span><span class="sxs-lookup"><span data-stu-id="0a411-114">The browser control loads the DOM and HTML body, and calls the event handler for the `window.onload` event.</span></span>

5. <span data-ttu-id="0a411-115">O aplicativo cliente Office carrega o ambiente de tempo de execução Office, que baixa e armazena em cache os arquivos da biblioteca da API JavaScript do servidor de rede de distribuição de conteúdo (CDN) e chama o manipulador de eventos do complemento para o evento [de inicialização](/javascript/api/office#office-initialize-reason-) do objeto [Office,](/javascript/api/office) se um manipulador tiver sido atribuído a ele.</span><span class="sxs-lookup"><span data-stu-id="0a411-115">The Office client application loads the runtime environment, which downloads and caches the Office JavaScript API library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it.</span></span> <span data-ttu-id="0a411-116">Neste momento, ele também verifica se algum retorno de chamada (ou `then()` funções encadeadas) foi autenticado (ou encadeado) para o `Office.onReady` identificador.</span><span class="sxs-lookup"><span data-stu-id="0a411-116">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="0a411-117">Para obter mais informações sobre a distinção `Office.initialize` entre e , consulte `Office.onReady` [Initialize your add-in](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="0a411-117">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

6. <span data-ttu-id="0a411-118">Quando o corpo de HTML e DOM terminar de carregar e o suplemento finalizar a inicialização, a função principal do suplemento poderá prosseguir.</span><span class="sxs-lookup"><span data-stu-id="0a411-118">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="0a411-119">Inicialização de um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="0a411-119">Startup of an Outlook add-in</span></span>

<span data-ttu-id="0a411-120">A figura a seguir mostra o fluxo de eventos envolvidos na inicialização de um suplemento do Outlook em execução no desktop, tablet ou smartphone.</span><span class="sxs-lookup"><span data-stu-id="0a411-120">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![Flow de eventos ao iniciar Outlook de um complemento.](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="0a411-122">Os eventos a seguir ocorrem quando um Outlook de usuário é iniciado.</span><span class="sxs-lookup"><span data-stu-id="0a411-122">The following events occur when an Outlook add-in starts.</span></span>

1. <span data-ttu-id="0a411-123">Quando é iniciado, o Outlook lê os manifestos XML para suplementos do Outlook que foram instalados na conta de email do usuário.</span><span class="sxs-lookup"><span data-stu-id="0a411-123">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>

2. <span data-ttu-id="0a411-124">O usuário seleciona um item no Outlook.</span><span class="sxs-lookup"><span data-stu-id="0a411-124">The user selects an item in Outlook.</span></span>

3. <span data-ttu-id="0a411-125">Se o item selecionado satisfizer as condições de ativação de um suplemento do Outlook, o Outlook ativará o suplemento e tornará seu botão visíveis na interface de usuário.</span><span class="sxs-lookup"><span data-stu-id="0a411-125">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>

4. <span data-ttu-id="0a411-p103">Se o usuário clicar no botão para iniciar o suplemento do Outlook, o Outlook abrirá a página HTML em um controle de navegador. As próximas duas etapas, as etapas 5 e 6, ocorrerem em paralelo.</span><span class="sxs-lookup"><span data-stu-id="0a411-p103">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>

5. <span data-ttu-id="0a411-128">O controle do navegador carrega o corpo DOM e HTML e chama o manipulador de eventos para o `onload` evento.</span><span class="sxs-lookup"><span data-stu-id="0a411-128">The browser control loads the DOM and HTML body, and calls the event handler for the `onload` event.</span></span>

6. <span data-ttu-id="0a411-129">O Outlook carrega o ambiente de tempo de execução, que baixa e armazena em cache a API do JavaScript para arquivos da biblioteca a partir do servidor da rede de distribuição de conteúdo (CDN) e chama manipulador de eventos para o evento [initialize](/javascript/api/office#office-initialize-reason-) do objeto do suplemento do [Office](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="0a411-129">Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it.</span></span> <span data-ttu-id="0a411-130">Neste momento, ele também verifica se algum retorno de chamada (ou `then()` funções encadeadas) foi autenticado (ou encadeado) para o `Office.onReady` identificador.</span><span class="sxs-lookup"><span data-stu-id="0a411-130">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="0a411-131">Para obter mais informações sobre a distinção `Office.initialize` entre e , consulte `Office.onReady` [Initialize your add-in](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="0a411-131">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

7. <span data-ttu-id="0a411-132">Quando o corpo de HTML e DOM terminar de carregar e o suplemento finalizar a inicialização, a função principal do suplemento poderá prosseguir.</span><span class="sxs-lookup"><span data-stu-id="0a411-132">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>

## <a name="see-also"></a><span data-ttu-id="0a411-133">Confira também</span><span class="sxs-lookup"><span data-stu-id="0a411-133">See also</span></span>

- [<span data-ttu-id="0a411-134">Entendendo a API de JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="0a411-134">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="0a411-135">Inicialize seu suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="0a411-135">Initialize your Office Add-in</span></span>](initialize-add-in.md)
