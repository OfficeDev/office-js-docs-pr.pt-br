# <a name="build-your-first-word-add-in"></a><span data-ttu-id="6bf5b-101">Crie o seu primeiro suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="6bf5b-101">Build your first Word add-in</span></span>

<span data-ttu-id="6bf5b-102">_Aplica-se a: Word 2016, Word para iPad, Word para Mac_</span><span class="sxs-lookup"><span data-stu-id="6bf5b-102">_Applies to: Word 2016, Word for iPad, Word for Mac_</span></span>

<span data-ttu-id="6bf5b-103">Neste artigo, você passará pelo processo de criar um suplemento do Word usando o jQuery e a API JavaScript para Word.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-103">In this article, you'll walk through the process of building a Word add-in by using jQuery and the Word JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="6bf5b-104">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="6bf5b-104">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="6bf5b-105">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="6bf5b-105">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="6bf5b-106">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="6bf5b-106">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="6bf5b-107">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="6bf5b-107">Create the add-in project</span></span>

1. <span data-ttu-id="6bf5b-108">Na barra de menus do Visual Studio, selecione **Arquivo**  >  **Novo**  >  **Projeto**.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="6bf5b-109">Na lista de tipos de projeto em **Visual C#** ou em **Visual Basic**, expanda a opção **Office/SharePoint**, escolha **Suplementos** e depois **Suplemento da Web do Word** como o tipo de projeto.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Word Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="6bf5b-110">Dê um nome ao projeto e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-110">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="6bf5b-p101">O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="6bf5b-113">Explorar a solução do Visual Studio</span><span class="sxs-lookup"><span data-stu-id="6bf5b-113">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="6bf5b-114">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="6bf5b-114">Update the code</span></span>

1. <span data-ttu-id="6bf5b-115">**Home.html** especifica o HTML que será renderizado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-115">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="6bf5b-116">Em **Home.html**, substitua o elemento `<body>` pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-116">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
    ```html
    <body>
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>    
        <div id="content-main">
            <div class="padding">
                <p>Choose the buttons below to add boilerplate text to the document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <br /><br />
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <br /><br />
                <button id="proverb">Add Chinese proverb</button>
            </div>
        </div>
        <br />
        <div id="supportedVersion"/>
    </body>
    ```

2. <span data-ttu-id="6bf5b-117">Abra o arquivo **Home.js** na raiz do projeto do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-117">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="6bf5b-118">Este arquivo especifica o script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-118">This file specifies the script for the add-in.</span></span> <span data-ttu-id="6bf5b-119">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-119">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';
    
    (function () {

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or later.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or later.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="6bf5b-120">Abra o arquivo **Home.css** na raiz do projeto do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-120">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="6bf5b-121">Este arquivo especifica os estilos personalizados para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-121">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="6bf5b-122">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-122">Replace the entire contents with the following code and save the file.</span></span>

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="6bf5b-123">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="6bf5b-123">Update the manifest</span></span>

1. <span data-ttu-id="6bf5b-124">Abra o arquivo de manifesto XML do projeto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-124">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="6bf5b-125">Este arquivo define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-125">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="6bf5b-126">O elemento `ProviderName` tem um valor de espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-126">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="6bf5b-127">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-127">Replace it with your name.</span></span>

3. <span data-ttu-id="6bf5b-128">O atributo `DefaultValue` do elemento `DisplayName` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-128">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="6bf5b-129">Substitua-o pelo **suplementos do My Office**.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-129">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="6bf5b-130">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-130">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="6bf5b-131">Substitua-o com **um suplemento do painel de tarefas do PowerPoint**.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-131">Replace it with **A task pane add-in for Word**.</span></span>

5. <span data-ttu-id="6bf5b-132">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-132">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="6bf5b-133">Experimente</span><span class="sxs-lookup"><span data-stu-id="6bf5b-133">Try it out</span></span>

1. <span data-ttu-id="6bf5b-p109">Usando o Visual Studio, teste o suplemento do Word recém-criado pressionando F5 ou escolhendo o botão **Iniciar** para abrir o Word com o botão de suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-p109">Using Visual Studio, test the newly created Word add-in by pressing F5 or choosing the **Start** button to launch Word with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="6bf5b-136">No Word, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-136">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="6bf5b-137">(Caso esteja usando a versão sem assinatura do Office 2016 ao invés da versão do Office 365, os botões personalizados não são compatíveis.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-137">(If you are using the non-subscription version of Office 2016, instead of the Office 365 version, then custom buttons are not supported.</span></span> <span data-ttu-id="6bf5b-138">Em vez disso, o painel de tarefas abrirá imediatamente.)</span><span class="sxs-lookup"><span data-stu-id="6bf5b-138">Instead, the task pane will open immediately.)</span></span>

    ![Uma captura de tela do Word com o botão Mostrar Painel de Tarefas realçado](../images/word-quickstart-addin-0.png)

3. <span data-ttu-id="6bf5b-140">No painel de tarefas, escolha qualquer um dos botões para adicionar o texto clichê ao documento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-140">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![Captura de tela do aplicativo Word com o suplemento de texto clichê carregado](../images/word-quickstart-addin-1b.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="6bf5b-142">Qualquer editor</span><span class="sxs-lookup"><span data-stu-id="6bf5b-142">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="6bf5b-143">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="6bf5b-143">Prerequisites</span></span>

- [<span data-ttu-id="6bf5b-144">Node.js</span><span class="sxs-lookup"><span data-stu-id="6bf5b-144">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="6bf5b-145">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-145">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-add-in-project"></a><span data-ttu-id="6bf5b-146">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="6bf5b-146">Create the add-in project</span></span>

1. <span data-ttu-id="6bf5b-147">Use o gerador Yeoman para criar um projeto de suplemento do Word.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-147">Use the Yeoman generator to create a Word add-in project.</span></span> <span data-ttu-id="6bf5b-148">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="6bf5b-148">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="6bf5b-149">**Escolha o tipo de projeto:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="6bf5b-149">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="6bf5b-150">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="6bf5b-150">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="6bf5b-151">**Qual será o nome do suplemento?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="6bf5b-151">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="6bf5b-152">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?:** `Word`</span><span class="sxs-lookup"><span data-stu-id="6bf5b-152">**Which Office client application would you like to support?:** `Word`</span></span>

    ![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-word-jquery.png)
    
    <span data-ttu-id="6bf5b-154">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-154">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="6bf5b-155">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-155">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

### <a name="update-the-code"></a><span data-ttu-id="6bf5b-156">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="6bf5b-156">Update the code</span></span>

1. <span data-ttu-id="6bf5b-157">No editor de código, abra **index.html** na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-157">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="6bf5b-158">Esse arquivo contém o HTML que será renderizado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-158">This file contains the HTML that will be rendered in the add-in's task pane.</span></span> 

2. <span data-ttu-id="6bf5b-159">Substitua o elemento `<body>` pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-159">Replace the `<body>` element with the following markup and save the file.</span></span>

    ```html
    <body>
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the buttons below to add boilerplate text to the document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <br /><br />
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <br /><br />
                <button id="proverb">Add Chinese proverb</button>
            </div>
        </div>
        <br />
        <div id="supportedVersion" />
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

2. <span data-ttu-id="6bf5b-160">Abra o arquivo **src/index.js** para especificar o script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-160">Open the file **src/index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="6bf5b-161">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-161">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="6bf5b-162">Esse script contém códigos de inicialização além do código que faz alterações no documento do Word inserindo texto no documento quando um botão é escolhido.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-162">This script contains initialization code as well as the code that makes changes to the Word document, by inserting text into the document when a button is chosen.</span></span> 

    ```js
    'use strict';
    
    (function () {

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or later.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or later.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="6bf5b-163">Abra o arquivo **app.css** na raiz do projeto para especificar os estilos personalizados do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-163">Open the file **app.css** in the root of the project to specify the custom styles for the add-in.</span></span> <span data-ttu-id="6bf5b-164">Substitua todo o conteúdo pelo que está a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-164">Replace the entire contents with the following and save the file.</span></span>

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="6bf5b-165">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="6bf5b-165">Update the manifest</span></span>

1. <span data-ttu-id="6bf5b-166">Abra o arquivo **manifest.xml** para definir as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-166">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="6bf5b-167">O elemento `ProviderName` tem um valor de espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-167">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="6bf5b-168">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-168">Replace it with your name.</span></span>

3. <span data-ttu-id="6bf5b-169">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-169">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="6bf5b-170">Substitua-o com **um suplemento do painel de tarefas do PowerPoint**.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-170">Replace it with **A task pane add-in for Word**.</span></span>

4. <span data-ttu-id="6bf5b-171">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-171">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="6bf5b-172">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="6bf5b-172">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="6bf5b-173">Experimente</span><span class="sxs-lookup"><span data-stu-id="6bf5b-173">Try it out</span></span>

1. <span data-ttu-id="6bf5b-174">Para realizar sideload do suplemento no Word, siga as instruções para a plataforma que você usará para executar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-174">To sideload the add-in within Word, follow the instructions for the platform you'll use to run your add-in.</span></span>

    - <span data-ttu-id="6bf5b-175">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="6bf5b-175">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="6bf5b-176">Word Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="6bf5b-176">Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="6bf5b-177">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="6bf5b-177">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="6bf5b-178">No Word, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-178">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Uma captura de tela do Word com o botão Mostrar painel de tarefas realçado](../images/word-quickstart-addin-2.png)

3. <span data-ttu-id="6bf5b-180">No painel de tarefas, escolha qualquer um dos botões para adicionar o texto clichê ao documento.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-180">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![Captura de tela do aplicativo Word com o suplemento de texto clichê carregado](../images/word-quickstart-addin-1.png)

---

## <a name="next-steps"></a><span data-ttu-id="6bf5b-182">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="6bf5b-182">Next steps</span></span>

<span data-ttu-id="6bf5b-183">Parabéns, você criou com êxito um suplemento do Word usando o jQuery!</span><span class="sxs-lookup"><span data-stu-id="6bf5b-183">Congratulations, you've successfully created a Word add-in using jQuery!</span></span> <span data-ttu-id="6bf5b-184">Em seguida, saiba mais sobre os recursos de um suplemento do Word e crie um suplemento mais complexo seguindo as etapas deste tutorial de suplemento do Word.</span><span class="sxs-lookup"><span data-stu-id="6bf5b-184">Next, learn more about the capabilities of a Word add-in and build a more complex add-in by following along with the Word add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="6bf5b-185">Tutorial de suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="6bf5b-185">Word add-in tutorial</span></span>](../tutorials/word-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="6bf5b-186">Confira também</span><span class="sxs-lookup"><span data-stu-id="6bf5b-186">See also</span></span>

* [<span data-ttu-id="6bf5b-187">Visão geral dos suplementos do Word</span><span class="sxs-lookup"><span data-stu-id="6bf5b-187">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
* <span data-ttu-id="6bf5b-188">
  [Exemplos de código do suplemento do Word](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,Word)</span><span class="sxs-lookup"><span data-stu-id="6bf5b-188">[Word add-in code samples](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,Word)</span></span>
* [<span data-ttu-id="6bf5b-189">Referências da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="6bf5b-189">Word JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview)
