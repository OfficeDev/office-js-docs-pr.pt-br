---
title: Crie seu primeiro suplemento do painel de tarefas do Word
description: Saiba como criar um suplemento do painel de tarefas do Word simples usando a API JS do Office.
ms.date: 07/07/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: ab8926eae6ddc41f82ab055d727b6279f357c316
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094468"
---
# <a name="build-your-first-word-task-pane-add-in"></a><span data-ttu-id="e6887-103">Crie seu primeiro suplemento do painel de tarefas do Word</span><span class="sxs-lookup"><span data-stu-id="e6887-103">Build your first Word task pane add-in</span></span>

<span data-ttu-id="e6887-104">_Aplica-se a: Word 2016 ou posterior no Windows, Word para iPad e Mac_</span><span class="sxs-lookup"><span data-stu-id="e6887-104">_Applies to: Word 2016 or later on Windows, and Word on iPad and Mac_</span></span>

<span data-ttu-id="e6887-105">Neste artigo, você aprenderá sobre o processo de criação de um suplemento do painel de tarefas do Word.</span><span class="sxs-lookup"><span data-stu-id="e6887-105">In this article, you'll walk through the process of building a Word task pane add-in.</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="e6887-106">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="e6887-106">Create the add-in</span></span>

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="yeoman-generator"></a>[<span data-ttu-id="e6887-107">Gerador do Yeoman</span><span class="sxs-lookup"><span data-stu-id="e6887-107">Yeoman generator</span></span>](#tab/yeomangenerator)

[!include[Redirect to the single sign-on (SSO) quick start](../includes/sso-quickstart-reference.md)]

### <a name="prerequisites"></a><span data-ttu-id="e6887-108">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="e6887-108">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="e6887-109">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="e6887-109">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="e6887-110">**Escolha o tipo de projeto:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="e6887-110">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="e6887-111">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="e6887-111">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="e6887-112">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="e6887-112">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="e6887-113">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="e6887-113">**Which Office client application would you like to support?**</span></span> `Word`

![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-word.png)

<span data-ttu-id="e6887-115">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="e6887-115">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="explore-the-project"></a><span data-ttu-id="e6887-116">Explore o projeto</span><span class="sxs-lookup"><span data-stu-id="e6887-116">Explore the project</span></span>

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-js.md)]

### <a name="try-it-out"></a><span data-ttu-id="e6887-117">Experimente</span><span class="sxs-lookup"><span data-stu-id="e6887-117">Try it out</span></span>

1. <span data-ttu-id="e6887-118">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="e6887-118">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="e6887-119">Conclua as etapas a seguir para iniciar o servidor Web local e fazer o sideload do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="e6887-119">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e6887-120">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="e6887-120">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="e6887-121">Se você for solicitado a instalar um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="e6887-121">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="e6887-122">Se você estiver testando o seu suplemento no Mac, execute o seguinte comando antes de continuar.</span><span class="sxs-lookup"><span data-stu-id="e6887-122">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="e6887-123">O servidor Web local é iniciado quando este comando é executado.</span><span class="sxs-lookup"><span data-stu-id="e6887-123">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="e6887-124">Para testar seu suplemento no Word, execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="e6887-124">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="e6887-125">Isso inicia o servidor Web local (caso ainda não esteja em execução) e abre o Word com o suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="e6887-125">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="e6887-126">Para testar seu suplemento no Word em um navegador, execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="e6887-126">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="e6887-127">Quando você executar este comando, o servidor Web local será iniciado (se ainda não estiver em execução).</span><span class="sxs-lookup"><span data-stu-id="e6887-127">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="e6887-128">Para usar o seu suplemento, abra um novo documento no Word na Web e em seguida realize o sideload no suplemento de acordo com as instruções em [Realizar Sideload nos Suplementos do Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="e6887-128">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="e6887-129">No Word, abra um novo documento, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="e6887-129">In Word, open a new document, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Uma captura de tela do aplicativo Word com o botão Mostrar Painel de Tarefas realçado](../images/word-quickstart-addin-2b.png)

4. <span data-ttu-id="e6887-131">Na parte inferior do painel de tarefas, escolha o link **Executar** para inserir o texto «Olá, Mundo» no documento com a fonte azul.</span><span class="sxs-lookup"><span data-stu-id="e6887-131">At the bottom of the task pane, choose the **Run** link to add the text "Hello World" to the document in blue font.</span></span>

    ![Captura de tela do aplicativo Word com o suplemento do painel de tarefas carregado](../images/word-quickstart-addin-1c.png)

### <a name="next-steps"></a><span data-ttu-id="e6887-133">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="e6887-133">Next steps</span></span>

<span data-ttu-id="e6887-134">Parabéns, você criou com êxito um suplemento do painel de tarefas do Word!</span><span class="sxs-lookup"><span data-stu-id="e6887-134">Congratulations, you've successfully created a Word task pane add-in!</span></span> <span data-ttu-id="e6887-135">Em seguida, saiba mais sobre os recursos de um suplemento do Word e crie um suplemento mais complexo seguindo as etapas deste [tutorial de suplemento do Word](../tutorials/word-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="e6887-135">Next, learn more about the capabilities of a Word add-in and build a more complex add-in by following along with the [Word add-in tutorial](../tutorials/word-tutorial.md).</span></span>

# <a name="visual-studio"></a>[<span data-ttu-id="e6887-136">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="e6887-136">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="e6887-137">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="e6887-137">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="e6887-138">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="e6887-138">Create the add-in project</span></span>


1. <span data-ttu-id="e6887-139">No Visual Studio, escolha **Criar um novo projeto**.</span><span class="sxs-lookup"><span data-stu-id="e6887-139">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="e6887-140">Usando a caixa de pesquisa, insira **suplemento**.</span><span class="sxs-lookup"><span data-stu-id="e6887-140">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="e6887-141">Escolha \*\*Suplemento do Word Web \*\*, em seguida, selecione **Próximo**.</span><span class="sxs-lookup"><span data-stu-id="e6887-141">Choose **Word Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="e6887-142">Nomeie seu projeto e selecione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="e6887-142">Name your project and select **Create**.</span></span>

4. <span data-ttu-id="e6887-143">Visual Studio creates a solution and its two projects appear in **Solution Explorer**.</span><span class="sxs-lookup"><span data-stu-id="e6887-143">Visual Studio creates a solution and its two projects appear in **Solution Explorer**.</span></span> <span data-ttu-id="e6887-144">The **Home.html** file opens in Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="e6887-144">The **Home.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="e6887-145">Explorar a solução do Visual Studio</span><span class="sxs-lookup"><span data-stu-id="e6887-145">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="e6887-146">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="e6887-146">Update the code</span></span>

1. <span data-ttu-id="e6887-147">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span><span class="sxs-lookup"><span data-stu-id="e6887-147">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="e6887-148">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span><span class="sxs-lookup"><span data-stu-id="e6887-148">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>

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

2. <span data-ttu-id="e6887-149">Open the file **Home.js** in the root of the web application project.</span><span class="sxs-lookup"><span data-stu-id="e6887-149">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="e6887-150">This file specifies the script for the add-in.</span><span class="sxs-lookup"><span data-stu-id="e6887-150">This file specifies the script for the add-in.</span></span> <span data-ttu-id="e6887-151">Replace the entire contents with the following code and save the file.</span><span class="sxs-lookup"><span data-stu-id="e6887-151">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', '1.1')) {
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
        });

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

3. <span data-ttu-id="e6887-152">Open the file **Home.css** in the root of the web application project.</span><span class="sxs-lookup"><span data-stu-id="e6887-152">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="e6887-153">This file specifies the custom styles for the add-in.</span><span class="sxs-lookup"><span data-stu-id="e6887-153">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="e6887-154">Replace the entire contents with the following code and save the file.</span><span class="sxs-lookup"><span data-stu-id="e6887-154">Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="e6887-155">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="e6887-155">Update the manifest</span></span>

1. <span data-ttu-id="e6887-156">Abra o arquivo de manifesto XML do projeto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="e6887-156">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="e6887-157">Este arquivo define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="e6887-157">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="e6887-158">The `ProviderName` element has a placeholder value.</span><span class="sxs-lookup"><span data-stu-id="e6887-158">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="e6887-159">Replace it with your name.</span><span class="sxs-lookup"><span data-stu-id="e6887-159">Replace it with your name.</span></span>

3. <span data-ttu-id="e6887-160">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span><span class="sxs-lookup"><span data-stu-id="e6887-160">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="e6887-161">Replace it with **My Office Add-in**.</span><span class="sxs-lookup"><span data-stu-id="e6887-161">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="e6887-162">The `DefaultValue` attribute of the `Description` element has a placeholder.</span><span class="sxs-lookup"><span data-stu-id="e6887-162">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="e6887-163">Replace it with **A task pane add-in for Word**.</span><span class="sxs-lookup"><span data-stu-id="e6887-163">Replace it with **A task pane add-in for Word**.</span></span>

5. <span data-ttu-id="e6887-164">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="e6887-164">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="e6887-165">Experimente</span><span class="sxs-lookup"><span data-stu-id="e6887-165">Try it out</span></span>

1. <span data-ttu-id="e6887-166">Using Visual Studio, test the newly created Word add-in by pressing **F5** or choosing the **Start** button to launch Word with the **Show Taskpane** add-in button displayed in the ribbon.</span><span class="sxs-lookup"><span data-stu-id="e6887-166">Using Visual Studio, test the newly created Word add-in by pressing **F5** or choosing the **Start** button to launch Word with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="e6887-167">The add-in will be hosted locally on IIS.</span><span class="sxs-lookup"><span data-stu-id="e6887-167">The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="e6887-168">No Word, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na Faixa de Opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="e6887-168">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="e6887-169">(Caso esteja usando uma versão de compra avulsa do Office, em vez da versão do Microsoft 365, os botões personalizados não serão compatíveis.</span><span class="sxs-lookup"><span data-stu-id="e6887-169">(If you are using the one-time purchase version of Office, instead of the Microsoft 365 version, then custom buttons are not supported.</span></span> <span data-ttu-id="e6887-170">Em vez disso, o painel de tarefas abrirá imediatamente.)</span><span class="sxs-lookup"><span data-stu-id="e6887-170">Instead, the task pane will open immediately.)</span></span>

    ![Uma captura de tela do Word com o botão Mostrar Painel de Tarefas realçado](../images/word-quickstart-addin-0.png)

3. <span data-ttu-id="e6887-172">No painel de tarefas, escolha qualquer um dos botões para adicionar o texto clichê ao documento.</span><span class="sxs-lookup"><span data-stu-id="e6887-172">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![Captura de tela do aplicativo Word com o suplemento de texto clichê carregado](../images/word-quickstart-addin-1b.png)

[!include[Console tool note](../includes/console-tool-note.md)]

### <a name="next-steps"></a><span data-ttu-id="e6887-174">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="e6887-174">Next steps</span></span>

<span data-ttu-id="e6887-175">Parabéns, você criou com êxito um suplemento do painel de tarefas do Word!</span><span class="sxs-lookup"><span data-stu-id="e6887-175">Congratulations, you've successfully created a Word task pane add-in!</span></span> <span data-ttu-id="e6887-176">Em seguida, saiba mais sobre como [desenvolver suplementos do Office com o Visual Studio](../develop/develop-add-ins-visual-studio.md).</span><span class="sxs-lookup"><span data-stu-id="e6887-176">Next, learn more about [developing Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

---

## <a name="see-also"></a><span data-ttu-id="e6887-177">Confira também</span><span class="sxs-lookup"><span data-stu-id="e6887-177">See also</span></span>

* [<span data-ttu-id="e6887-178">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e6887-178">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="e6887-179">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="e6887-179">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="e6887-180">Desenvolver Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="e6887-180">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="e6887-181">Visão geral dos suplementos do Word</span><span class="sxs-lookup"><span data-stu-id="e6887-181">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
* [<span data-ttu-id="e6887-182">Exemplos de código do suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="e6887-182">Word add-in code samples</span></span>](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,Word)
* [<span data-ttu-id="e6887-183">Referências da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="e6887-183">Word JavaScript API reference</span></span>](../reference/overview/word-add-ins-reference-overview.md)
