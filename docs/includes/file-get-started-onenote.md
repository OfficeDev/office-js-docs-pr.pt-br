# <a name="build-your-first-onenote-add-in"></a><span data-ttu-id="72a8c-101">Criar seu primeiro suplemento do OneNote</span><span class="sxs-lookup"><span data-stu-id="72a8c-101">Build your first OneNote add-in</span></span>

<span data-ttu-id="72a8c-102">Neste artigo, você passará pelo processo de criar um suplemento do OneNote usando o jQuery e a API JavaScript para Office.</span><span class="sxs-lookup"><span data-stu-id="72a8c-102">In this article, you'll walk through the process of building a OneNote add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="72a8c-103">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="72a8c-103">Prerequisites</span></span>

- [<span data-ttu-id="72a8c-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="72a8c-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="72a8c-105">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="72a8c-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a><span data-ttu-id="72a8c-106">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="72a8c-106">Create the add-in project</span></span>

1. <span data-ttu-id="72a8c-107">Use o gerador Yeoman para criar um projeto de suplemento do OneNote.</span><span class="sxs-lookup"><span data-stu-id="72a8c-107">Use the Yeoman generator to create a OneNote add-in project.</span></span> <span data-ttu-id="72a8c-108">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="72a8c-108">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="72a8c-109">**Escolha o tipo de projeto:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="72a8c-109">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="72a8c-110">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="72a8c-110">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="72a8c-111">**Qual será o nome do suplemento?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="72a8c-111">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="72a8c-112">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?:** `Onenote`</span><span class="sxs-lookup"><span data-stu-id="72a8c-112">**Which Office client application would you like to support?:** `Onenote`</span></span>

    ![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-onenote-jquery.png)
    
    <span data-ttu-id="72a8c-114">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="72a8c-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="72a8c-115">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="72a8c-115">Navigate to the root folder of the project in the Terminal app, and from Terminal run:</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="72a8c-116">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="72a8c-116">Update the code</span></span>

1. <span data-ttu-id="72a8c-117">No editor de código, abra **index.html** na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="72a8c-117">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="72a8c-118">Esse arquivo contém o HTML que será renderizado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="72a8c-118">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="72a8c-119">Substitua o elemento `<body>` pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="72a8c-119">Replace the `<body>` element inside the  element with the following markup and save the file.</span></span> 

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

3. <span data-ttu-id="72a8c-120">Abra o arquivo **src\index.js** para especificar o script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="72a8c-120">Open the file **src\index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="72a8c-121">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="72a8c-121">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    import * as OfficeHelpers from "@microsoft/office-js-helpers";

    Office.initialize = (reason) => {
        $(document).ready(() => {
            $('#addOutline').click(addOutlineToPage);
        });
    };
    
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

4. <span data-ttu-id="72a8c-122">Abra o arquivo **app.css** para especificar os estilos personalizados do suplemento.</span><span class="sxs-lookup"><span data-stu-id="72a8c-122">Open the file **app.css** to specify the custom styles for the add-in.</span></span> <span data-ttu-id="72a8c-123">Substitua todo o conteúdo pelo que está a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="72a8c-123">Replace the entire contents with the following and save the file.</span></span>

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

## <a name="update-the-manifest"></a><span data-ttu-id="72a8c-124">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="72a8c-124">Update the manifest</span></span>

1. <span data-ttu-id="72a8c-125">Abra o arquivo **manifest.xml** para definir as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="72a8c-125">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="72a8c-126">O elemento `ProviderName` tem um valor de espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="72a8c-126">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="72a8c-127">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="72a8c-127">Replace it with your name.</span></span>

3. <span data-ttu-id="72a8c-128">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="72a8c-128">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="72a8c-129">Substitua-o por **um suplemento do painel de tarefas do OneNote**.</span><span class="sxs-lookup"><span data-stu-id="72a8c-129">Replace it with **A task pane add-in for OneNote**.</span></span>

4. <span data-ttu-id="72a8c-130">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="72a8c-130">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="72a8c-131">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="72a8c-131">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a><span data-ttu-id="72a8c-132">Experimente</span><span class="sxs-lookup"><span data-stu-id="72a8c-132">Try it out</span></span>

1. <span data-ttu-id="72a8c-133">No [OneNote Online](https://www.onenote.com/notebooks), abra um bloco de anotações.</span><span class="sxs-lookup"><span data-stu-id="72a8c-133">In [OneNote Online](https://www.onenote.com/notebooks), open a notebook.</span></span>

2. <span data-ttu-id="72a8c-134">Escolha **Inserir > Suplementos do Office** para abrir a caixa de diálogo Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="72a8c-134">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="72a8c-135">Se você estiver conectado à sua conta de consumidor, selecione a guia **MEUS SUPLEMENTOS** e escolha  **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="72a8c-135">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="72a8c-136">Se você estiver conectado à sua conta corporativa ou de estudante, selecione a guia **MINHA ORGANIZAÇÃO** e escolha  **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="72a8c-136">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="72a8c-137">A imagem a seguir mostra a guia **MEUS SUPLEMENTOS** para blocos de anotações do consumidor.</span><span class="sxs-lookup"><span data-stu-id="72a8c-137">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. <span data-ttu-id="72a8c-138">Na caixa de diálogo Carregar Suplemento, navegue até **manifest.xml** na pasta do projeto e escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="72a8c-138">In the Upload Add-in dialog, browse to **one-note-add-in-manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

4. <span data-ttu-id="72a8c-139">Na guia **Página Inicial**, na faixa de opções, escolha o botão **Mostrar Painel de Tarefas**.</span><span class="sxs-lookup"><span data-stu-id="72a8c-139">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="72a8c-140">O painel de tarefa do suplemento abre em um iFrame ao lado da página do OneNote.</span><span class="sxs-lookup"><span data-stu-id="72a8c-140">6- The add-in opens in an iFrame next to the OneNote page.</span></span>

5. <span data-ttu-id="72a8c-141">Insira este conteúdo HTML na área de texto e escolha **Adicionar estrutura de tópicos**.</span><span class="sxs-lookup"><span data-stu-id="72a8c-141">Enter some text in the text area and then choose **Add outline**.</span></span>  

    ```html
    <ol>
    <li>Item #1</li>
    <li>Item #2</li>
    <li>Item #3</li>
    <li>Item #4</li>
    </ol>
    ```

    <span data-ttu-id="72a8c-142">A estrutura de tópicos especificada é adicionada à página.</span><span class="sxs-lookup"><span data-stu-id="72a8c-142">The outline that you specified is added to the page.</span></span>

    ![O suplemento do OneNote criado a partir deste passo a passo](../images/onenote-first-add-in-3.png)

## <a name="troubleshooting-and-tips"></a><span data-ttu-id="72a8c-144">Dicas e solução de problemas</span><span class="sxs-lookup"><span data-stu-id="72a8c-144">Troubleshooting and tips</span></span>

- <span data-ttu-id="72a8c-p108">Você pode depurar o suplemento usando as ferramentas de desenvolvedor do seu navegador. Quando você estiver usando o servidor Web Gulp e depurando no Internet Explore ou no Chrome, você pode salvar as alterações localmente e apenas atualize o iFrame do suplemento.</span><span class="sxs-lookup"><span data-stu-id="72a8c-p108">You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.</span></span>

- <span data-ttu-id="72a8c-p109">Quando você inspecionar um objeto do OneNote, as propriedades que estão atualmente disponíveis usam valores reais de exibição. As propriedades que precisam ser carregadas exibem *undefined*. Expanda o nó `_proto_` para ver as propriedades definidas no objeto, mas que ainda não foram carregadas.</span><span class="sxs-lookup"><span data-stu-id="72a8c-p109">When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `_proto_` node to see properties that are defined on the object but are not yet loaded.</span></span>

   ![Carregar o objeto do OneNote em um depurador](../images/onenote-debug.png)

- <span data-ttu-id="72a8c-p110">Você precisa habilitar conteúdo misto no navegador, se o seu suplemento usar todos os recursos HTTP. Os suplementos de produção devem usar apenas recursos HTTPS seguros.</span><span class="sxs-lookup"><span data-stu-id="72a8c-p110">You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.</span></span>

- <span data-ttu-id="72a8c-153">É possível abrir os suplementos do Painel de Tarefas em praticamente qualquer lugar, mas os suplementos de conteúdo podem ser inseridos apenas no conteúdo normal da página (ou seja, fora títulos, imagens, iFrames, etc.).</span><span class="sxs-lookup"><span data-stu-id="72a8c-153">Task pane add-ins can be opened from anywhere, but content add-ins can only be inserted inside regular page content (i.e. not in titles, images, iFrames, etc.).</span></span> 

## <a name="next-steps"></a><span data-ttu-id="72a8c-154">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="72a8c-154">Next steps</span></span>

<span data-ttu-id="72a8c-155">Parabéns, você criou com êxito um suplemento do OneNote!</span><span class="sxs-lookup"><span data-stu-id="72a8c-155">Congratulations, you've successfully created a OneNote add-in!</span></span> <span data-ttu-id="72a8c-156">Em seguida, saiba mais sobre os principais conceitos de criação de suplementos do OneNote.</span><span class="sxs-lookup"><span data-stu-id="72a8c-156">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="72a8c-157">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="72a8c-157">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="72a8c-158">Confira também</span><span class="sxs-lookup"><span data-stu-id="72a8c-158">See also</span></span>

- [<span data-ttu-id="72a8c-159">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="72a8c-159">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="72a8c-160">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="72a8c-160">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference?view=office-js)
- [<span data-ttu-id="72a8c-161">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="72a8c-161">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="72a8c-162">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="72a8c-162">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
