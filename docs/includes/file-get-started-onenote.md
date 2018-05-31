# <a name="build-your-first-onenote-add-in"></a><span data-ttu-id="517d1-101">Criar seu primeiro suplemento do OneNote</span><span class="sxs-lookup"><span data-stu-id="517d1-101">Build your first OneNote add-in</span></span>

<span data-ttu-id="517d1-102">Neste artigo, você passará pelo processo de criar um suplemento do OneNote usando o jQuery e a API JavaScript para Office.</span><span class="sxs-lookup"><span data-stu-id="517d1-102">In this article, you'll walk through the process of building a OneNote add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="517d1-103">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="517d1-103">Prerequisites</span></span>

- [<span data-ttu-id="517d1-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="517d1-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="517d1-105">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="517d1-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a><span data-ttu-id="517d1-106">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="517d1-106">Create the add-in project</span></span>

1. <span data-ttu-id="517d1-107">Crie uma pasta na sua unidade local e nomeie-a como `my-onenote-addin`.</span><span class="sxs-lookup"><span data-stu-id="517d1-107">Create a folder on your local drive and name it `my-onenote-addin`.</span></span> <span data-ttu-id="517d1-108">Esse é o local em que você criará os arquivos para seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="517d1-108">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="517d1-109">Navegue até a nova pasta.</span><span class="sxs-lookup"><span data-stu-id="517d1-109">Navigate to your new folder.</span></span>

    ```bash
    cd my-onenote-addin
    ```

3. <span data-ttu-id="517d1-110">Use o gerador Yeoman para criar um projeto de suplemento do OneNote.</span><span class="sxs-lookup"><span data-stu-id="517d1-110">Use the Yeoman generator to create a OneNote add-in project.</span></span> <span data-ttu-id="517d1-111">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="517d1-111">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="517d1-112">**Gostaria de criar uma nova subpasta para o seu projeto?:** `No`</span><span class="sxs-lookup"><span data-stu-id="517d1-112">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="517d1-113">**Como deseja nomear seu suplemento?:** `OneNote Add-in`</span><span class="sxs-lookup"><span data-stu-id="517d1-113">**What do you want to name your add-in?:** `OneNote Add-in`</span></span>
    - <span data-ttu-id="517d1-114">**Para qual aplicativo cliente do Office você deseja suporte?:** `OneNote`</span><span class="sxs-lookup"><span data-stu-id="517d1-114">**Which Office client application would you like to support?:** `OneNote`</span></span>
    - <span data-ttu-id="517d1-115">**Gostaria de criar um novo suplemento?:** `Yes`</span><span class="sxs-lookup"><span data-stu-id="517d1-115">**Would you like to create a new add-in?:** `Yes`</span></span>
    - <span data-ttu-id="517d1-116">**Gostaria de usar o TypeScript?:** `No`</span><span class="sxs-lookup"><span data-stu-id="517d1-116">**Would you like to use TypeScript?:** `No`</span></span>
    - <span data-ttu-id="517d1-117">**Escolha a estrutura:** `Jquery`</span><span class="sxs-lookup"><span data-stu-id="517d1-117">**Choose a framework:** `Jquery`</span></span>

    <span data-ttu-id="517d1-p103">O gerador perguntará se você deseja abrir **resource.html**. Não é necessário abri-lo para este tutorial, mas fique à vontade em fazer isso se tiver curiosidade. Escolha Sim ou Não para concluir o assistente e deixar o gerador fazer seu trabalho.</span><span class="sxs-lookup"><span data-stu-id="517d1-p103">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-onenote-jquery.png)


## <a name="update-the-code"></a><span data-ttu-id="517d1-122">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="517d1-122">Update the code</span></span>

1. <span data-ttu-id="517d1-123">No editor de código, abra **index.html** na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="517d1-123">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="517d1-124">Esse arquivo contém o HTML que será renderizado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="517d1-124">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="517d1-125">Substitua o elemento `<main>` dentro do elemento `<body>` com a marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="517d1-125">Replace the `<main>` element inside the `<body>` element with the following markup and save the file.</span></span> <span data-ttu-id="517d1-126">Isso adiciona uma área de texto e um botão usando [componentes do Office UI Fabric](http://dev.office.com/fabric/components).</span><span class="sxs-lookup"><span data-stu-id="517d1-126">This adds a text area and a button using [Office UI Fabric components](http://dev.office.com/fabric/components).</span></span>

    ```html
    <main class="ms-welcome__main">
        <br />
        <p class="ms-font-l">Enter content below</p>
        <div class="ms-TextField ms-TextField--placeholder">
            <textarea id="textBox" rows="5"></textarea>
        </div>
        <button id="addOutline" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
            <span class="ms-Button-label">Add Outline</span>
            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
            <span class="ms-Button-description">Adds the content above to the current page.</span>
        </button>
    </main>
    ```

3. <span data-ttu-id="517d1-127">Abra o arquivo **app.js** para especificar o script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="517d1-127">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="517d1-128">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="517d1-128">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        Office.initialize = function (reason) {
            $(document).ready(function () {
                app.initialize();

                // Set up event handler for the UI.
                $('#addOutline').click(addOutlineToPage);
            });
        };

        // Add the contents of the text area to the page.
        function addOutlineToPage() {        
            OneNote.run(function (context) {
                var html = '<p>' + $('#textBox').val() + '</p>';

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.             
                page.load('title'); 

                // Add an outline with the specified HTML to the page.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log('Added outline to page ' + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error); 
                        console.log("Error: " + error); 
                        if (error instanceof OfficeExtension.Error) { 
                            console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                        } 
                    }); 
            });
        }
    })();
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="517d1-129">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="517d1-129">Update the manifest</span></span>

1. <span data-ttu-id="517d1-130">Abra o arquivo **one-note-add-in-manifest.xml** para definir as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="517d1-130">Open the file **one-note-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="517d1-131">O elemento `ProviderName` tem um valor de espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="517d1-131">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="517d1-132">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="517d1-132">Replace it with your name.</span></span>

3. <span data-ttu-id="517d1-133">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="517d1-133">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="517d1-134">Substitua-o por **um suplemento do painel de tarefas do OneNote**.</span><span class="sxs-lookup"><span data-stu-id="517d1-134">Replace it with **A task pane add-in for OneNote**.</span></span>

4. <span data-ttu-id="517d1-135">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="517d1-135">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="OneNote Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="517d1-136">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="517d1-136">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a><span data-ttu-id="517d1-137">Experimente</span><span class="sxs-lookup"><span data-stu-id="517d1-137">Try it out</span></span>

1. <span data-ttu-id="517d1-138">No [OneNote Online](https://www.onenote.com/notebooks), abra um bloco de anotações.</span><span class="sxs-lookup"><span data-stu-id="517d1-138">In [OneNote Online](https://www.onenote.com/notebooks), open a notebook.</span></span>

2. <span data-ttu-id="517d1-139">Escolha **Inserir > Suplementos do Office** para abrir a caixa de diálogo Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="517d1-139">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="517d1-140">Se você estiver conectado à sua conta de consumidor, selecione a guia **MEUS SUPLEMENTOS** e escolha  **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="517d1-140">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="517d1-141">Se você estiver conectado à sua conta corporativa ou de estudante, selecione a guia **MINHA ORGANIZAÇÃO** e escolha  **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="517d1-141">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="517d1-142">A imagem a seguir mostra a guia **MEUS SUPLEMENTOS** para blocos de anotações do consumidor.</span><span class="sxs-lookup"><span data-stu-id="517d1-142">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. <span data-ttu-id="517d1-143">Na caixa de diálogo Carregar suplemento, navegue até **one-note-add-in-manifest.xml** na pasta do projeto e escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="517d1-143">In the Upload Add-in dialog, browse to **one-note-add-in-manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

4. <span data-ttu-id="517d1-144">Na guia **Página Inicial**, escolha o botão **Exibir painel de tarefas** na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="517d1-144">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="517d1-145">O painel de tarefas do suplemento abre em um iFrame perto da página do OneNote.</span><span class="sxs-lookup"><span data-stu-id="517d1-145">6- The add-in opens in an iFrame next to the OneNote page.</span></span>

5. <span data-ttu-id="517d1-146">Insira algum texto na área de texto e escolha **Adicionar estrutura de tópicos**.</span><span class="sxs-lookup"><span data-stu-id="517d1-146">Enter some text in the text area and then choose **Add outline**.</span></span> <span data-ttu-id="517d1-147">O texto inserido é adicionado à pagina.</span><span class="sxs-lookup"><span data-stu-id="517d1-147">The text you entered is added to the page.</span></span> 

    ![O suplemento do OneNote criado a partir deste passo a passo](../images/onenote-first-add-in.png)

## <a name="troubleshooting-and-tips"></a><span data-ttu-id="517d1-149">Dicas e solução de problemas</span><span class="sxs-lookup"><span data-stu-id="517d1-149">Troubleshooting and tips</span></span>

- <span data-ttu-id="517d1-p111">Você pode depurar o suplemento usando as ferramentas de desenvolvedor do seu navegador. Quando você estiver usando o servidor Web Gulp e depurando no Internet Explore ou no Chrome, você pode salvar as alterações localmente e apenas atualize o iFrame do suplemento.</span><span class="sxs-lookup"><span data-stu-id="517d1-p111">You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.</span></span>

- <span data-ttu-id="517d1-p112">Quando você inspecionar um objeto do OneNote, as propriedades que estão atualmente disponíveis usam valores reais de exibição. As propriedades que precisam ser carregadas exibem *undefined*. Expanda o nó `_proto_` para ver as propriedades definidas no objeto, mas que ainda não foram carregadas.</span><span class="sxs-lookup"><span data-stu-id="517d1-p112">When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `_proto_` node to see properties that are defined on the object but are not yet loaded.</span></span>

   ![Carregar o objeto do OneNote em um depurador](../images/onenote-debug.png)

- <span data-ttu-id="517d1-p113">Você precisa habilitar conteúdo misto no navegador, se o seu suplemento usar todos os recursos HTTP. Os suplementos de produção devem usar apenas recursos HTTPS seguros.</span><span class="sxs-lookup"><span data-stu-id="517d1-p113">You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.</span></span>

- <span data-ttu-id="517d1-158">É possível abrir os suplementos do Painel de Tarefas em praticamente qualquer lugar, mas os suplementos de conteúdo podem ser inseridos apenas no conteúdo normal da página (ou seja, fora títulos, imagens, iFrames, etc.).</span><span class="sxs-lookup"><span data-stu-id="517d1-158">Task pane add-ins can be opened from anywhere, but content add-ins can only be inserted inside regular page content (i.e. not in titles, images, iFrames, etc.).</span></span> 

## <a name="next-steps"></a><span data-ttu-id="517d1-159">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="517d1-159">Next steps</span></span>

<span data-ttu-id="517d1-160">Parabéns, você criou com êxito um suplemento do OneNote!</span><span class="sxs-lookup"><span data-stu-id="517d1-160">Congratulations, you've successfully created a OneNote add-in!</span></span> <span data-ttu-id="517d1-161">Em seguida, saiba mais sobre os principais conceitos de criação de suplementos do OneNote.</span><span class="sxs-lookup"><span data-stu-id="517d1-161">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="517d1-162">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="517d1-162">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="517d1-163">Veja também</span><span class="sxs-lookup"><span data-stu-id="517d1-163">See also</span></span>

- [<span data-ttu-id="517d1-164">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="517d1-164">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="517d1-165">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="517d1-165">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="517d1-166">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="517d1-166">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="517d1-167">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="517d1-167">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
