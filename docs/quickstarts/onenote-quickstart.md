---
title: Criar o seu primeiro suplemento do OneNote
description: ''
ms.date: 03/19/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 378d691d1994a2d22166afc5338007400f7a48af
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871917"
---
# <a name="build-your-first-onenote-add-in"></a><span data-ttu-id="a8b31-102">Criar o seu primeiro suplemento do OneNote</span><span class="sxs-lookup"><span data-stu-id="a8b31-102">Build your first OneNote add-in</span></span>

<span data-ttu-id="a8b31-103">Neste artigo, você passará pelo processo de criar um suplemento do OneNote usando o jQuery e a API JavaScript para Office.</span><span class="sxs-lookup"><span data-stu-id="a8b31-103">In this article, you'll walk through the process of building a OneNote add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a8b31-104">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="a8b31-104">Prerequisites</span></span>

- [<span data-ttu-id="a8b31-105">Node.js</span><span class="sxs-lookup"><span data-stu-id="a8b31-105">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="a8b31-106">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="a8b31-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a><span data-ttu-id="a8b31-107">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="a8b31-107">Create the add-in project</span></span>

1. <span data-ttu-id="a8b31-108">Use o gerador Yeoman para criar um projeto de suplemento do OneNote.</span><span class="sxs-lookup"><span data-stu-id="a8b31-108">Use the Yeoman generator to create a OneNote add-in project.</span></span> <span data-ttu-id="a8b31-109">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="a8b31-109">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="a8b31-110">**Escolha o tipo de projeto:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="a8b31-110">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="a8b31-111">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="a8b31-111">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="a8b31-112">**Qual será o nome do suplemento?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="a8b31-112">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="a8b31-113">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?:** `Onenote`</span><span class="sxs-lookup"><span data-stu-id="a8b31-113">**Which Office client application would you like to support?:** `Onenote`</span></span>

    ![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-onenote-jquery.png)
    
    <span data-ttu-id="a8b31-115">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="a8b31-115">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="a8b31-116">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="a8b31-116">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="a8b31-117">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="a8b31-117">Update the code</span></span>

1. <span data-ttu-id="a8b31-p102">No editor de código, abra **index.html** na raiz do projeto. Esse arquivo contém o HTML que será renderizado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="a8b31-p102">In your code editor, open **index.html** in the root of the project. This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="a8b31-120">Substitua o elemento `<body>` pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a8b31-120">Replace the `<body>` element with the following markup and save the file.</span></span> 

    ```html
    <body class="ms-font-m ms-welcome">
        <header class="ms-welcome__header ms-bgColor-themeDark ms-u-fadeIn500">
            <h2 class="ms-fontSize-xxl ms-fontWeight-regular ms-fontColor-white">OneNote Add-in</h1>
        </header>
        <main id="app-body" class="ms-welcome__main">
            <br />
            <p class="ms-font-m">Enter HTML content here:</p>
            <div class="ms-TextField ms-TextField--placeholder">
                <textarea id="textBox" rows="8" cols="30"></textarea>
            </div>
            <button id="addOutline" class="ms-Button ms-Button--primary">
                <span class="ms-Button-label">Add outline</span>
            </button>
        </main>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. <span data-ttu-id="a8b31-121">Abra o arquivo **src\index.js** para especificar o script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="a8b31-121">Open the file **src\index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="a8b31-122">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a8b31-122">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    import * as OfficeHelpers from "@microsoft/office-js-helpers";

    Office.onReady(() => {
        // Office is ready
        $(document).ready(() => {
            // The document is ready
            $('#addOutline').click(addOutlineToPage);
        });
    });
    
    async function addOutlineToPage() {
        try {
            await OneNote.run(async context => {
                var html = "<p>" + $("#textBox").val() + "</p>";

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.
                page.load("title");

                // Add text to the page by using the specified HTML.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log("Added outline to page " + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error);
                        console.log("Error: " + error);
                        if (error instanceof OfficeExtension.Error) {
                            console.log("Debug info: " + JSON.stringify(error.debugInfo));
                        }
                    });
                });
        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }
    ```

4. <span data-ttu-id="a8b31-123">Abra o arquivo **app.css** para especificar os estilos personalizados do suplemento.</span><span class="sxs-lookup"><span data-stu-id="a8b31-123">Open the file **app.css** to specify the custom styles for the add-in.</span></span> <span data-ttu-id="a8b31-124">Substitua todo o conteúdo pelo que está a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a8b31-124">Replace the entire contents with the following and save the file.</span></span>

    ```css
    html, body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
    }

    ul, p, h1, h2, h3, h4, h5, h6 {
        margin: 0;
        padding: 0;
    }

    .ms-welcome {
        position: relative;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        min-height: 500px;
        min-width: 320px;
        overflow: auto;
        overflow-x: hidden;
    }

    .ms-welcome__header {
        min-height: 30px;
        padding: 0px;
        padding-bottom: 5px;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: center;
        -webkit-justify-content: flex-end;
        justify-content: flex-end;
    }

    .ms-welcome__header > h1 {
        margin-top: 5px;
        text-align: center;
    }

    .ms-welcome__main {
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: left;
        -webkit-flex: 1 0 0;
        flex: 1 0 0;
        padding: 30px 20px;
    }

    .ms-welcome__main > h2 {
        width: 100%;
        text-align: left;
    }

    @media (min-width: 0) and (max-width: 350px) {
        .ms-welcome__features {
            width: 100%;
        }
    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="a8b31-125">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="a8b31-125">Update the manifest</span></span>

1. <span data-ttu-id="a8b31-126">Abra o arquivo **manifest.xml** para definir as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="a8b31-126">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="a8b31-p105">O elemento `ProviderName` tem um valor de espaço reservado. Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="a8b31-p105">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="a8b31-p106">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado. Substitua-o por **um suplemento do painel de tarefas do OneNote**.</span><span class="sxs-lookup"><span data-stu-id="a8b31-p106">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for OneNote**.</span></span>

4. <span data-ttu-id="a8b31-131">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a8b31-131">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="a8b31-132">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="a8b31-132">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a><span data-ttu-id="a8b31-133">Experimente</span><span class="sxs-lookup"><span data-stu-id="a8b31-133">Try it out</span></span>

1. <span data-ttu-id="a8b31-134">No [OneNote Online](https://www.onenote.com/notebooks), abra um bloco de anotações.</span><span class="sxs-lookup"><span data-stu-id="a8b31-134">In [OneNote Online](https://www.onenote.com/notebooks), open a notebook.</span></span>

2. <span data-ttu-id="a8b31-135">Escolha **Inserir > Suplementos do Office** para abrir a caixa de diálogo Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="a8b31-135">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="a8b31-136">Se você estiver conectado à sua conta de consumidor, selecione a guia **MEUS SUPLEMENTOS** e escolha  **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="a8b31-136">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="a8b31-137">Se você estiver conectado à sua conta corporativa ou de estudante, selecione a guia **MINHA ORGANIZAÇÃO** e escolha  **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="a8b31-137">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="a8b31-138">A imagem a seguir mostra a guia **MEUS SUPLEMENTOS** para blocos de anotações do consumidor.</span><span class="sxs-lookup"><span data-stu-id="a8b31-138">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. <span data-ttu-id="a8b31-139">Na caixa de diálogo Carregar Suplemento, navegue até **manifest.xml** na pasta do projeto e escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="a8b31-139">In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

4. <span data-ttu-id="a8b31-140">Na guia **Página Inicial**, na faixa de opções, escolha o botão **Mostrar Painel de Tarefas**.</span><span class="sxs-lookup"><span data-stu-id="a8b31-140">From the **Home** tab, choose the **Show Taskpane** button in the ribbon.</span></span> <span data-ttu-id="a8b31-141">O painel de tarefa do suplemento abre em um iFrame ao lado da página do OneNote.</span><span class="sxs-lookup"><span data-stu-id="a8b31-141">The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

5. <span data-ttu-id="a8b31-142">Insira este conteúdo HTML na área de texto e escolha **Adicionar estrutura de tópicos**.</span><span class="sxs-lookup"><span data-stu-id="a8b31-142">Enter the following HTML content in the text area, and then choose **Add outline**.</span></span>  

    ```html
    <ol>
    <li>Item #1</li>
    <li>Item #2</li>
    <li>Item #3</li>
    <li>Item #4</li>
    </ol>
    ```

    <span data-ttu-id="a8b31-143">A estrutura de tópicos especificada é adicionada à página.</span><span class="sxs-lookup"><span data-stu-id="a8b31-143">The outline that you specified is added to the page.</span></span>

    ![O suplemento do OneNote criado a partir deste passo a passo](../images/onenote-first-add-in-3.png)

## <a name="troubleshooting-and-tips"></a><span data-ttu-id="a8b31-145">Dicas e solução de problemas</span><span class="sxs-lookup"><span data-stu-id="a8b31-145">Troubleshooting and tips</span></span>

- <span data-ttu-id="a8b31-p108">Você pode depurar o suplemento usando as ferramentas de desenvolvedor do seu navegador. Quando você estiver usando o servidor Web Gulp e depurando no Internet Explore ou no Chrome, você pode salvar as alterações localmente e apenas atualize o iFrame do suplemento.</span><span class="sxs-lookup"><span data-stu-id="a8b31-p108">You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.</span></span>

- <span data-ttu-id="a8b31-p109">Quando você inspecionar um objeto do OneNote, as propriedades que estão atualmente disponíveis usam valores reais de exibição. As propriedades que precisam ser carregadas exibem *undefined*. Expanda o nó `_proto_` para ver as propriedades definidas no objeto, mas que ainda não foram carregadas.</span><span class="sxs-lookup"><span data-stu-id="a8b31-p109">When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `_proto_` node to see properties that are defined on the object but are not yet loaded.</span></span>

   ![Carregar o objeto do OneNote em um depurador](../images/onenote-debug.png)

- <span data-ttu-id="a8b31-p110">Você precisa habilitar conteúdo misto no navegador, se o seu suplemento usar todos os recursos HTTP. Os suplementos de produção devem usar apenas recursos HTTPS seguros.</span><span class="sxs-lookup"><span data-stu-id="a8b31-p110">You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.</span></span>

- <span data-ttu-id="a8b31-154">É possível abrir os suplementos do Painel de Tarefas em praticamente qualquer lugar, mas os suplementos de conteúdo podem ser inseridos apenas no conteúdo normal da página (ou seja, fora títulos, imagens, iFrames, etc.).</span><span class="sxs-lookup"><span data-stu-id="a8b31-154">Task pane add-ins can be opened from anywhere, but content add-ins can only be inserted inside regular page content (i.e. not in titles, images, iFrames, etc.).</span></span> 

## <a name="next-steps"></a><span data-ttu-id="a8b31-155">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="a8b31-155">Next steps</span></span>

<span data-ttu-id="a8b31-p111">Parabéns, você criou com êxito um suplemento do OneNote! Em seguida, saiba mais sobre os principais conceitos de criação de suplementos do OneNote.</span><span class="sxs-lookup"><span data-stu-id="a8b31-p111">Congratulations, you've successfully created a OneNote add-in! Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="a8b31-158">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="a8b31-158">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="a8b31-159">Confira também</span><span class="sxs-lookup"><span data-stu-id="a8b31-159">See also</span></span>

- [<span data-ttu-id="a8b31-160">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="a8b31-160">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="a8b31-161">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="a8b31-161">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="a8b31-162">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="a8b31-162">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="a8b31-163">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="a8b31-163">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

