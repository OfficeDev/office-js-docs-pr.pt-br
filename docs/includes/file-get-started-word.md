# <a name="build-your-first-word-add-in"></a><span data-ttu-id="07397-101">Compilar seu primeiro suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="07397-101">Build your first Word add-in</span></span>

<span data-ttu-id="07397-102">_Aplica-se a: Word 2016, Word para iPad, Word para Mac_</span><span class="sxs-lookup"><span data-stu-id="07397-102">_Applies to: Word 2016, Word for iPad, Word for Mac_</span></span>

<span data-ttu-id="07397-103">Neste artigo, voc? passar? pelo processo de criar um suplemento do Word usando o jQuery e a API JavaScript para Word.</span><span class="sxs-lookup"><span data-stu-id="07397-103">In this article, you'll walk through the process of building a Word add-in by using jQuery and the Word JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="07397-104">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="07397-104">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="07397-105">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="07397-105">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="07397-106">Pr?-requisitos</span><span class="sxs-lookup"><span data-stu-id="07397-106">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="07397-107">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="07397-107">Create the add-in project</span></span>

1. <span data-ttu-id="07397-108">Na barra de menus do Visual Studio, selecione **Arquivo**  >  **Novo**  >  **Projeto**.</span><span class="sxs-lookup"><span data-stu-id="07397-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="07397-109">Na lista de tipos de projeto em **Visual C#** ou em **Visual Basic**, expanda a op??o **Office/SharePoint**, escolha **Suplementos** e depois **Suplemento da Web do Word** como o tipo de projeto.</span><span class="sxs-lookup"><span data-stu-id="07397-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Word Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="07397-110">D? um nome ao projeto e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="07397-110">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="07397-p101">O Visual Studio cria uma solu??o, e os dois projetos dele s?o exibidos no **Gerenciador de Solu??es**. O arquivo **Home.html** ? aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="07397-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="07397-113">Explorar a solu??o do Visual Studio</span><span class="sxs-lookup"><span data-stu-id="07397-113">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="07397-114">Atualizar o c?digo</span><span class="sxs-lookup"><span data-stu-id="07397-114">Update the code</span></span>

1. <span data-ttu-id="07397-115">**Home.html** especifica o HTML que ser? renderizado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="07397-115">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="07397-116">Em **Home.html**, substitua o elemento `<body>` pela marca??o a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="07397-116">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
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

2. <span data-ttu-id="07397-117">Abra o arquivo **Home.js** na raiz do projeto do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="07397-117">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="07397-118">Este arquivo especifica o script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="07397-118">This file specifies the script for the add-in.</span></span> <span data-ttu-id="07397-119">Substitua todo o conte?do pelo c?digo a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="07397-119">Replace the entire contents with the following code and save the file.</span></span>

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
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
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

3. <span data-ttu-id="07397-120">Abra o arquivo **Home.css** na raiz do projeto do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="07397-120">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="07397-121">Este arquivo especifica os estilos personalizados para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="07397-121">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="07397-122">Substitua todo o conte?do pelo c?digo a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="07397-122">Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="07397-123">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="07397-123">Update the manifest</span></span>

1. <span data-ttu-id="07397-124">Abra o arquivo de manifesto XML do projeto do Suplemento.</span><span class="sxs-lookup"><span data-stu-id="07397-124">Open the XML manifest file in the Add-in project.</span></span> <span data-ttu-id="07397-125">Este arquivo define as configura??es e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="07397-125">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="07397-126">O elemento `ProviderName` tem um valor de espa?o reservado.</span><span class="sxs-lookup"><span data-stu-id="07397-126">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="07397-127">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="07397-127">Replace it with your name.</span></span>

3. <span data-ttu-id="07397-128">O atributo `DefaultValue` do elemento `DisplayName` tem um espa?o reservado.</span><span class="sxs-lookup"><span data-stu-id="07397-128">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="07397-129">Substitua-o pelo **suplementos do My Office**.</span><span class="sxs-lookup"><span data-stu-id="07397-129">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="07397-130">O atributo `DefaultValue` do elemento `Description` tem um espa?o reservado.</span><span class="sxs-lookup"><span data-stu-id="07397-130">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="07397-131">Substitua-o com **um suplemento do painel de tarefas do PowerPoint**.</span><span class="sxs-lookup"><span data-stu-id="07397-131">Replace it with **A task pane add-in for Word**.</span></span>

5. <span data-ttu-id="07397-132">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="07397-132">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="07397-133">Experimente</span><span class="sxs-lookup"><span data-stu-id="07397-133">Try it out</span></span>

1. <span data-ttu-id="07397-p109">Usando o Visual Studio, teste o suplemento do Word rec?m-criado pressionando F5 ou escolhendo o bot?o **Iniciar** para abrir o Word com o bot?o de suplemento **Mostrar painel de tarefas** exibido na faixa de op??es. O suplemento ser? hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="07397-p109">Using Visual Studio, test the newly created Word add-in by pressing F5 or choosing the **Start** button to launch Word with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="07397-136">No Word, escolha a guia **P?gina Inicial** e o bot?o **Mostrar Painel de Tarefas** na faixa de op??es para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="07397-136">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Uma captura de tela do Word com o bot?o Mostrar painel de tarefas real?ado](../images/word-quickstart-addin-0.png)

3. <span data-ttu-id="07397-138">No painel de tarefas, escolha qualquer um dos bot?es para adicionar o texto clich? ao documento.</span><span class="sxs-lookup"><span data-stu-id="07397-138">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![Captura de tela do aplicativo Word com o suplemento de texto clich? carregado](../images/word-quickstart-addin-1b.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="07397-140">Qualquer editor</span><span class="sxs-lookup"><span data-stu-id="07397-140">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="07397-141">Pr?-requisitos</span><span class="sxs-lookup"><span data-stu-id="07397-141">Prerequisites</span></span>

- [<span data-ttu-id="07397-142">Node.js</span><span class="sxs-lookup"><span data-stu-id="07397-142">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="07397-143">Instale a ?ltima vers?o do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="07397-143">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-add-in-project"></a><span data-ttu-id="07397-144">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="07397-144">Create the add-in project</span></span>

1. <span data-ttu-id="07397-145">Crie uma pasta na sua unidade local e nomeie-a como `my-word-addin`.</span><span class="sxs-lookup"><span data-stu-id="07397-145">Create a folder on your local drive and name it `my-word-addin`.</span></span> <span data-ttu-id="07397-146">Esse ? o local em que voc? criar? os arquivos para seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="07397-146">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="07397-147">Navegue at? a nova pasta.</span><span class="sxs-lookup"><span data-stu-id="07397-147">Navigate to your new folder.</span></span>

    ```bash
    cd my-word-addin
    ```

3. <span data-ttu-id="07397-148">Use o gerador Yeoman para criar um projeto do suplemento do Word.</span><span class="sxs-lookup"><span data-stu-id="07397-148">Use the Yeoman generator to create a Word add-in project.</span></span> <span data-ttu-id="07397-149">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="07397-149">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="07397-150">**Gostaria de criar uma nova subpasta para o seu projeto?** `No`</span><span class="sxs-lookup"><span data-stu-id="07397-150">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="07397-151">**Como deseja nomear seu suplemento?** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="07397-151">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="07397-152">**Para qual aplicativo cliente do Office voc? deseja suporte?** `Word`</span><span class="sxs-lookup"><span data-stu-id="07397-152">**Which Office client application would you like to support?:** `Word`</span></span>
    - <span data-ttu-id="07397-153">**Gostaria de criar um novo suplemento?** `Yes`</span><span class="sxs-lookup"><span data-stu-id="07397-153">**Would you like to create a new add-in?:** `Yes`</span></span>
    - <span data-ttu-id="07397-154">**Gostaria de usar o TypeScript?** `No`</span><span class="sxs-lookup"><span data-stu-id="07397-154">**Would you like to use TypeScript?:** `No`</span></span>
    - <span data-ttu-id="07397-155">**Escolha a estrutura:** `Jquery`</span><span class="sxs-lookup"><span data-stu-id="07397-155">**Choose a framework:** `Jquery`</span></span>

    <span data-ttu-id="07397-p112">O gerador perguntar? se voc? deseja abrir **resource.html**. N?o ? necess?rio abri-lo para este tutorial, mas fique ? vontade em fazer isso se tiver curiosidade. Escolha Sim ou N?o para concluir o assistente e deixar o gerador fazer seu trabalho.</span><span class="sxs-lookup"><span data-stu-id="07397-p112">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-word-jquery.png)

### <a name="update-the-code"></a><span data-ttu-id="07397-160">Atualizar o c?digo</span><span class="sxs-lookup"><span data-stu-id="07397-160">Update the code</span></span>

1. <span data-ttu-id="07397-161">No editor de c?digo, abra **index.html** na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="07397-161">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="07397-162">Esse arquivo cont?m o HTML que ser? renderizado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="07397-162">This file contains the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="07397-163">Substitua todo o conte?do pelo c?digo a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="07397-163">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="07397-164">Este suplemento exibir? tr?s bot?es, e quando qualquer um dos bot?es for escolhido, o texto clich? ser? adicionado ao documento.</span><span class="sxs-lookup"><span data-stu-id="07397-164">This add-in will display three buttons and when any of the buttons are chosen, boilerplate text will be added to the document.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <title>Boilerplate text app</title>
            <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
            <script src="app.js" type="text/javascript"></script>
            <link href="app.css" rel="stylesheet" type="text/css" />
        </head>
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
    </html>
    ```

2. <span data-ttu-id="07397-165">Abra o arquivo **app.js** para especificar o script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="07397-165">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="07397-166">Substitua todo o conte?do pelo c?digo a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="07397-166">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="07397-167">Esse script cont?m c?digos de inicializa??o al?m do c?digo que faz altera??es no documento do Word inserindo texto no documento quando um bot?o ? escolhido.</span><span class="sxs-lookup"><span data-stu-id="07397-167">This script contains initialization code as well as the code that makes changes to the Word document, by inserting text into the document when a button is chosen.</span></span> 

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
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
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

3. <span data-ttu-id="07397-168">Abra o arquivo **app.css** na raiz do projeto para especificar os estilos personalizados do suplemento.</span><span class="sxs-lookup"><span data-stu-id="07397-168">Open the file **app.css** in the root of the project to specify the custom styles for the add-in.</span></span> <span data-ttu-id="07397-169">Substitua todo o conte?do pelo que est? a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="07397-169">Replace the entire contents with the following and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="07397-170">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="07397-170">Update the manifest</span></span>

1. <span data-ttu-id="07397-171">Abra o arquivo **my-office-add-in-manifest.xml** para definir as configura??es e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="07397-171">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="07397-172">O elemento `ProviderName` tem um valor de espa?o reservado.</span><span class="sxs-lookup"><span data-stu-id="07397-172">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="07397-173">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="07397-173">Replace it with your name.</span></span>

3. <span data-ttu-id="07397-174">O atributo `DefaultValue` do elemento `Description` tem um espa?o reservado.</span><span class="sxs-lookup"><span data-stu-id="07397-174">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="07397-175">Substitua-o com **um suplemento do painel de tarefas do PowerPoint**.</span><span class="sxs-lookup"><span data-stu-id="07397-175">Replace it with **A task pane add-in for Word**.</span></span>

4. <span data-ttu-id="07397-176">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="07397-176">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="07397-177">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="07397-177">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="07397-178">Experimente</span><span class="sxs-lookup"><span data-stu-id="07397-178">Try it out</span></span>

1. <span data-ttu-id="07397-179">Para realizar sideload do suplemento no Word, siga as instru??es para a plataforma que voc? usar? para executar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="07397-179">To sideload the add-in within Word, follow the instructions for the platform you'll use to run your add-in.</span></span>

    - <span data-ttu-id="07397-180">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="07397-180">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="07397-181">Word Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="07397-181">Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="07397-182">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="07397-182">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="07397-183">No Word, escolha a guia **P?gina Inicial** e o bot?o **Mostrar Painel de Tarefas** na faixa de op??es para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="07397-183">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Uma captura de tela do Word com o bot?o Mostrar painel de tarefas real?ado](../images/word-quickstart-addin-2.png)

3. <span data-ttu-id="07397-185">No painel de tarefas, escolha qualquer um dos bot?es para adicionar o texto clich? ao documento.</span><span class="sxs-lookup"><span data-stu-id="07397-185">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![Captura de tela do aplicativo Word com o suplemento de texto clich? carregado](../images/word-quickstart-addin-1.png)

---

## <a name="next-steps"></a><span data-ttu-id="07397-187">Pr?ximas etapas</span><span class="sxs-lookup"><span data-stu-id="07397-187">Next steps</span></span>

<span data-ttu-id="07397-188">Parab?ns, voc? criou com ?xito um suplemento do Word usando o jQuery!</span><span class="sxs-lookup"><span data-stu-id="07397-188">Congratulations, you've successfully created a Word add-in using jQuery!</span></span> <span data-ttu-id="07397-189">Em seguida, saiba mais sobre os recursos de um suplemento do Word e crie um suplemento mais complexo seguindo as etapas deste tutorial de suplemento do Word.</span><span class="sxs-lookup"><span data-stu-id="07397-189">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="07397-190">Tutorial do suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="07397-190">Word add-in tutorial</span></span>](../tutorials/word-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="07397-191">Veja tamb?m</span><span class="sxs-lookup"><span data-stu-id="07397-191">See also</span></span>

* [<span data-ttu-id="07397-192">Vis?o geral dos suplementos do Word</span><span class="sxs-lookup"><span data-stu-id="07397-192">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
* [<span data-ttu-id="07397-193">Exemplos de c?digo do suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="07397-193">Word add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=word,office%20add-ins)
* [<span data-ttu-id="07397-194">Refer?ncias da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="07397-194">Word JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)
