---
title: Carregar o ambiente de tempo de execução e DOM
description: Carregar o ambiente de tempo de execução de suplementos do Office e DOM
ms.date: 07/01/2019
localization_priority: Normal
ms.openlocfilehash: 2ea5f1fdc42fe1ffde30f8145fd0c24599c7e702
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718913"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="412b7-103">Carregar o ambiente de tempo de execução e DOM</span><span class="sxs-lookup"><span data-stu-id="412b7-103">Loading the DOM and runtime environment</span></span>

<span data-ttu-id="412b7-104">Um suplemento deve garantir que o DOM e o ambiente de tempo de execução de Suplementos do Office sejam carregados antes de executar sua própria lógica personalizada.</span><span class="sxs-lookup"><span data-stu-id="412b7-104">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span>

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="412b7-105">Inicialização de um suplemento de conteúdo ou de painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="412b7-105">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="412b7-106">A figura a seguir mostra o fluxo de eventos envolvidos na inicialização de um suplemento de conteúdo ou de painel de tarefas no Excel, no PowerPoint, no Project ou no Word.</span><span class="sxs-lookup"><span data-stu-id="412b7-106">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, or Word.</span></span>

![Fluxo de eventos ao iniciar um suplemento de conteúdo ou de painel de tarefas](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="412b7-108">Os eventos a seguir ocorrem quando um suplemento de conteúdo ou de painel de tarefas é iniciado:</span><span class="sxs-lookup"><span data-stu-id="412b7-108">The following events occur when a content or task pane add-in starts:</span></span>

1. <span data-ttu-id="412b7-109">O usuário abre um documento que já contém um suplemento ou insere um suplemento no documento.</span><span class="sxs-lookup"><span data-stu-id="412b7-109">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>

2. <span data-ttu-id="412b7-110">O aplicativo host do Office lê o manifesto XML do suplemento do AppSource, de um catálogo de aplicativos no SharePoint ou do catálogo de pastas compartilhadas de onde ele se origina.</span><span class="sxs-lookup"><span data-stu-id="412b7-110">The Office host application reads the add-in's XML manifest from AppSource, an app catalog on SharePoint, or the shared folder catalog it originates from.</span></span>

3. <span data-ttu-id="412b7-111">O aplicativo host do Office abre a página de HTML do suplemento em um controle de navegador.</span><span class="sxs-lookup"><span data-stu-id="412b7-111">The Office host application opens the add-in's HTML page in a browser control.</span></span>

    <span data-ttu-id="412b7-p101">As próximas duas etapas, as etapas 4 e 5, ocorrem de forma assíncrona e em paralelo. Por esse motivo, o código do suplemento deve garantir que o DOM e o ambiente do tempo de execução do suplemento tenham terminado de carregar antes de prosseguir.</span><span class="sxs-lookup"><span data-stu-id="412b7-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>

4. <span data-ttu-id="412b7-114">O controle do navegador carrega o corpo do HTML e DOM e chama o manipulador de eventos `window.onload` para o evento.</span><span class="sxs-lookup"><span data-stu-id="412b7-114">The browser control loads the DOM and HTML body, and calls the event handler for the `window.onload` event.</span></span>

5. <span data-ttu-id="412b7-115">O aplicativo host do Office carrega o ambiente de tempo de execução, que baixa e armazena em cache a API do JavaScript para arquivos da biblioteca a partir do servidor da rede de distribuição de conteúdo (CDN) e chama manipulador de eventos do suplemento para o evento [initialize](/javascript/api/office#office-initialize-reason-) do objeto [Office](/javascript/api/office), se um identificador for atribuído a ele.</span><span class="sxs-lookup"><span data-stu-id="412b7-115">The Office host application loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it.</span></span> <span data-ttu-id="412b7-116">Neste momento, ele também verifica se algum retorno de chamada (ou `then()` funções encadeadas) foi autenticado (ou encadeado) para o `Office.onReady` identificador.</span><span class="sxs-lookup"><span data-stu-id="412b7-116">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="412b7-117">Para obter mais informações sobre a distinção `Office.initialize` entre `Office.onReady`o e o, consulte [Initialize Your Add-in](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="412b7-117">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

6. <span data-ttu-id="412b7-118">Quando o corpo de HTML e DOM terminar de carregar e o suplemento finalizar a inicialização, a função principal do suplemento poderá prosseguir.</span><span class="sxs-lookup"><span data-stu-id="412b7-118">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="412b7-119">Inicialização de um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="412b7-119">Startup of an Outlook add-in</span></span>

<span data-ttu-id="412b7-120">A figura a seguir mostra o fluxo de eventos envolvidos na inicialização de um suplemento do Outlook em execução no desktop, tablet ou smartphone.</span><span class="sxs-lookup"><span data-stu-id="412b7-120">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![Fluxo de eventos ao inicializar um suplemento do Outlook](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="412b7-122">Os eventos a seguir ocorrem quando um suplemento Outlook é iniciado:</span><span class="sxs-lookup"><span data-stu-id="412b7-122">The following events occur when an Outlook add-in starts:</span></span>

1. <span data-ttu-id="412b7-123">Quando é iniciado, o Outlook lê os manifestos XML para suplementos do Outlook que foram instalados na conta de email do usuário.</span><span class="sxs-lookup"><span data-stu-id="412b7-123">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>

2. <span data-ttu-id="412b7-124">O usuário seleciona um item no Outlook.</span><span class="sxs-lookup"><span data-stu-id="412b7-124">The user selects an item in Outlook.</span></span>

3. <span data-ttu-id="412b7-125">Se o item selecionado satisfizer as condições de ativação de um suplemento do Outlook, o Outlook ativará o suplemento e tornará seu botão visíveis na interface de usuário.</span><span class="sxs-lookup"><span data-stu-id="412b7-125">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>

4. <span data-ttu-id="412b7-p103">Se o usuário clicar no botão para iniciar o suplemento do Outlook, o Outlook abrirá a página HTML em um controle de navegador. As próximas duas etapas, as etapas 5 e 6, ocorrerem em paralelo.</span><span class="sxs-lookup"><span data-stu-id="412b7-p103">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>

5. <span data-ttu-id="412b7-128">O controle do navegador carrega o corpo do HTML e DOM e chama o manipulador de eventos `onload` para o evento.</span><span class="sxs-lookup"><span data-stu-id="412b7-128">The browser control loads the DOM and HTML body, and calls the event handler for the `onload` event.</span></span>

6. <span data-ttu-id="412b7-129">O Outlook carrega o ambiente de tempo de execução, que baixa e armazena em cache a API do JavaScript para arquivos da biblioteca a partir do servidor da rede de distribuição de conteúdo (CDN) e chama manipulador de eventos para o evento [initialize](/javascript/api/office#office-initialize-reason-) do objeto do suplemento do [Office](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="412b7-129">Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it.</span></span> <span data-ttu-id="412b7-130">Neste momento, ele também verifica se algum retorno de chamada (ou `then()` funções encadeadas) foi autenticado (ou encadeado) para o `Office.onReady` identificador.</span><span class="sxs-lookup"><span data-stu-id="412b7-130">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="412b7-131">Para obter mais informações sobre a distinção `Office.initialize` entre `Office.onReady`o e o, consulte [Initialize Your Add-in](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="412b7-131">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

7. <span data-ttu-id="412b7-132">Quando o corpo de HTML e DOM terminar de carregar e o suplemento finalizar a inicialização, a função principal do suplemento poderá prosseguir.</span><span class="sxs-lookup"><span data-stu-id="412b7-132">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="checking-the-load-status"></a><span data-ttu-id="412b7-133">Verificar o status de carregamento</span><span class="sxs-lookup"><span data-stu-id="412b7-133">Checking the load status</span></span>

<span data-ttu-id="412b7-134">Uma maneira de verificar se o ambiente de tempo de execução e o DOM concluíram o carregamento é usar a função [.ready()](https://api.jquery.com/ready/) do jQuery: `$(document).ready()`.</span><span class="sxs-lookup"><span data-stu-id="412b7-134">One way to check that both the DOM and the runtime environment have finished loading is to use the jQuery [.ready()](https://api.jquery.com/ready/) function: `$(document).ready()`.</span></span> <span data-ttu-id="412b7-135">Por exemplo, o manipulador `onReady` de eventos a seguir garante que o dom seja carregado primeiro antes que o código específico para inicializar o suplemento seja executado.</span><span class="sxs-lookup"><span data-stu-id="412b7-135">For example, the following `onReady` event handler makes sure the DOM is first loaded before the code specific to initializing the add-in runs.</span></span> <span data-ttu-id="412b7-136">Subsequentemente, `onReady` o manipulador continua a usar a propriedade [Mailbox. Item](/javascript/api/outlook/office.mailbox) para obter o item atualmente selecionado no Outlook e chama a função principal do suplemento, `initDialer`.</span><span class="sxs-lookup"><span data-stu-id="412b7-136">Subsequently, the `onReady` handler proceeds to use the [mailbox.item](/javascript/api/outlook/office.mailbox) property to obtain the currently selected item in Outlook, and calls the main function of the add-in, `initDialer`.</span></span>

```js
Office.onReady()
    .then(
        // Checks for the DOM to load.
        $(document).ready(function () {
            // After the DOM is loaded, add-in-specific code can run.
            var mailbox = Office.context.mailbox;
            _Item = mailbox.item;
            initDialer();
        });
);
```

<span data-ttu-id="412b7-137">Como alternativa, você pode usar o mesmo código em um `initialize` manipulador de eventos, conforme mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="412b7-137">Alternatively, you can use the same code in an `initialize` event handler as shown in the following example.</span></span>

```js
Office.initialize = function () {
    // Checks for the DOM to load.
    $(document).ready(function () {
        // After the DOM is loaded, add-in-specific code can run.
        var mailbox = Office.context.mailbox;
        _Item = mailbox.item;
        initDialer();
    });
}
```

<span data-ttu-id="412b7-138">Essa mesma técnica pode ser usada nos `onReady` manipuladores `initialize` ou de qualquer suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="412b7-138">This same technique can be used in the `onReady` or `initialize` handlers of any Office Add-in.</span></span>

<span data-ttu-id="412b7-139">O suplemento do Outlook de amostra de discagem telefônica mostra uma abordagem ligeiramente diferente usando somente o JavaScript para verificar essas mesmas condições.</span><span class="sxs-lookup"><span data-stu-id="412b7-139">The phone dialer sample Outlook add-in shows a slightly different approach using only JavaScript to check these same conditions.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="412b7-140">Mesmo que seu suplemento não tenha tarefas de inicialização para executar, você deve incluir pelo menos uma chamada ou atribuir `Office.onReady` uma função de `Office.initialize` manipulador de eventos mínima, conforme mostrado nos exemplos a seguir.</span><span class="sxs-lookup"><span data-stu-id="412b7-140">Even if your add-in has no initialization tasks to perform, you must include at least a call of `Office.onReady` or assign minimal `Office.initialize` event handler function as shown in the following examples.</span></span>
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> <span data-ttu-id="412b7-141">Se você não chamar `Office.onReady` ou atribuir um `Office.initialize` manipulador de eventos, seu suplemento poderá gerar um erro quando for iniciado.</span><span class="sxs-lookup"><span data-stu-id="412b7-141">If you do not call `Office.onReady` or assign an `Office.initialize` event handler, your add-in may raise an error when it starts.</span></span> <span data-ttu-id="412b7-142">Além disso, se um usuário tentar usar o suplemento com um cliente virtual do Office Online, como Excel, PowerPoint ou Outlook, ele não funcionará.</span><span class="sxs-lookup"><span data-stu-id="412b7-142">Also, if a user attempts to use your add-in with an Office web client, such as Excel, PowerPoint, or Outlook, it will fail to run.</span></span>
>
> <span data-ttu-id="412b7-143">Se o suplemento incluir mais de uma página, sempre que carregar uma nova página, a página deverá chamar `Office.onReady` ou atribuir um manipulador de `Office.initialize` eventos.</span><span class="sxs-lookup"><span data-stu-id="412b7-143">If your add-in includes more than one page, whenever it loads a new page that page must either call `Office.onReady` or assign an `Office.initialize` event handler.</span></span>

## <a name="see-also"></a><span data-ttu-id="412b7-144">Também confira</span><span class="sxs-lookup"><span data-stu-id="412b7-144">See also</span></span>

- [<span data-ttu-id="412b7-145">Entendendo a API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="412b7-145">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="412b7-146">Inicialize seu suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="412b7-146">Initialize your Office Add-in</span></span>](initialize-add-in.md)
