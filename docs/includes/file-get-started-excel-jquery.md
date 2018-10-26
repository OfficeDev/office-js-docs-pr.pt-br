# <a name="build-an-excel-add-in-using-jquery"></a><span data-ttu-id="81d3f-101">Criar um suplemento do Excel usando o jQuery</span><span class="sxs-lookup"><span data-stu-id="81d3f-101">Build an Excel add-in using jQuery</span></span>

<span data-ttu-id="81d3f-102">Neste artigo, você passará pelo processo de criação de um suplemento do Excel usando o jQuery e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="81d3f-102">In this article, you'll walk through the process of building an Excel add-in by using jQuery and the Excel JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="81d3f-103">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="81d3f-103">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="81d3f-104">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="81d3f-104">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="81d3f-105">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="81d3f-105">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="81d3f-106">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="81d3f-106">Create the add-in project</span></span>

1. <span data-ttu-id="81d3f-107">Na barra de menus do Visual Studio, selecione **Arquivo** >  **Novo** > **Projeto**.</span><span class="sxs-lookup"><span data-stu-id="81d3f-107">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="81d3f-108">Na lista de tipos de projeto em **Visual C#** ou em **Visual Basic**, expanda a opção **Office/SharePoint**, escolha **Suplementos** e depois **Suplemento da Web do Excel** como o tipo de projeto.</span><span class="sxs-lookup"><span data-stu-id="81d3f-108">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="81d3f-109">Dê um nome ao projeto e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="81d3f-109">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="81d3f-110">Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel** e clique em **Concluir** para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="81d3f-110">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="81d3f-p101">O Visual Studio cria uma solução e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="81d3f-113">Explorar a solução do Visual Studio</span><span class="sxs-lookup"><span data-stu-id="81d3f-113">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="81d3f-114">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="81d3f-114">Update the code</span></span>

1. <span data-ttu-id="81d3f-p102">**Home.HTML** especifica o HTML que será processado no painel de tarefas do suplemento. Em **Home.html**, substitua o elemento `<body>`  com a seguinte marcação e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p102">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the button below to set the color of the selected range to green.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="set-color">Set color</button>
            </div>
        </div>
    </body>
    ```

2. <span data-ttu-id="81d3f-p103">Abra o arquivo **Home.js** na raiz do projeto de aplicativo da web. Este arquivo especifica o script para o suplemento. Substitua todo o conteúdo pelo seguinte código e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p103">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

    ```js
    'use strict';

    (function () {
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="81d3f-p104">Abra o arquivo **Home.css** na raiz do projeto de aplicativo da web. Este arquivo especifica os estilos personalizados para o suplemento. Substitua todo o conteúdo pelo seguinte código e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p104">Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

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

### <a name="update-the-manifest"></a><span data-ttu-id="81d3f-123">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="81d3f-123">Update the manifest</span></span>

1. <span data-ttu-id="81d3f-p105">Abra o arquivo de manifesto XML no projeto de suplemento. Este arquivo define as configurações e recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p105">Open the XML manifest file in the add-in project. This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="81d3f-p106">O elemento `ProviderName` tem um valor de espaço reservado. Substitua-o pelo seu nome.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p106">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="81d3f-p107">O atributo `DefaultValue` do elemento `DisplayName` tem um espaço reservado. Substitua-o pelo **Meu suplemento do Office**.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p107">The `DefaultValue` attribute of the `DisplayName` element has a placeholder. Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="81d3f-p108">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado. Substitua-o por **Um suplemento do painel de tarefas para o Excel**.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p108">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="81d3f-132">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="81d3f-132">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="81d3f-133">Experimente</span><span class="sxs-lookup"><span data-stu-id="81d3f-133">Try it out</span></span>

1. <span data-ttu-id="81d3f-p109">Usando o Visual Studio, teste o suplemento do Excel recém-criado pressionando F5 ou clicando no botão **Iniciar** para abrir o Excel com o botão de suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p109">Using Visual Studio, test the newly created Excel add-in by pressing F5 or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="81d3f-136">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="81d3f-136">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="81d3f-138">Selecione qualquer intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="81d3f-138">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="81d3f-139">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="81d3f-139">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="81d3f-141">Qualquer editor</span><span class="sxs-lookup"><span data-stu-id="81d3f-141">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="81d3f-142">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="81d3f-142">Prerequisites</span></span>

- [<span data-ttu-id="81d3f-143">Node.js</span><span class="sxs-lookup"><span data-stu-id="81d3f-143">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="81d3f-144">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador Yeoman para suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="81d3f-144">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a><span data-ttu-id="81d3f-145">Criar o aplicativo Web</span><span class="sxs-lookup"><span data-stu-id="81d3f-145">Create the web app</span></span>

1. <span data-ttu-id="81d3f-146">Use o gerador Yeoman para criar um projeto de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="81d3f-146">Use the Yeoman generator to create an Outlook add-in project.</span></span> <span data-ttu-id="81d3f-147">Execute o comando a seguir e responda às mensagens da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="81d3f-147">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="81d3f-148">**Escolha um tipo de projeto:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="81d3f-148">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="81d3f-149">**Escolha um tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="81d3f-149">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="81d3f-150">**Como deseja nomear seu suplemento?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="81d3f-150">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="81d3f-151">**Qual aplicativo cliente do Office você gostaria de suportar?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="81d3f-151">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Gerador do Yeoman](../images/yo-office-jquery.png)
    
    <span data-ttu-id="81d3f-153">Depois de concluir o assistente, o gerador criará o projeto e instalará os componentes de suporte do Node.</span><span class="sxs-lookup"><span data-stu-id="81d3f-153">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="81d3f-154">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="81d3f-154">Navigate to the root folder of the web application project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

### <a name="update-the-code"></a><span data-ttu-id="81d3f-155">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="81d3f-155">Update the code</span></span> 

1. <span data-ttu-id="81d3f-p111">No editor de código, abra o **index.html** na raiz do projeto. Este arquivo especifica o HTML que será processado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p111">In your code editor, open **index.html** in the root of the project. This file specifies the HTML that will be rendered in the add-in's task pane.</span></span> 
 
2. <span data-ttu-id="81d3f-158">Dentro do **index.html**, substitua a tag `body` pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="81d3f-158">Within **index.html**, replace the generated `body` tag with the following markup, and save the file.</span></span>
 
    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the button below to set the color of the selected range to green.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="set-color">Set color</button>
            </div>
        </div>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>    
    ```

3. <span data-ttu-id="81d3f-p112">Abra o arquivo **src/index.js** para especificar o script do suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p112">Open the file **src\index.js** to specify the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';
    
    (function () {
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

4. <span data-ttu-id="81d3f-p113">Abra o arquivo **app.css** para especificar os estilos personalizados para o suplemento. Substitua todo o conteúdo com o código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p113">Open the file **app.css** to specify the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="81d3f-163">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="81d3f-163">Update the manifest</span></span>

1. <span data-ttu-id="81d3f-164">Abra o arquivo **manifest.xml** para definir as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="81d3f-164">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="81d3f-165">O elemento `ProviderName` tem um valor de espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="81d3f-165">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="81d3f-166">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="81d3f-166">Replace it with your name.</span></span>

3. <span data-ttu-id="81d3f-p115">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado. Substitua-o por **Um suplemento do painel de tarefas para o Excel**.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p115">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="81d3f-169">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="81d3f-169">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="81d3f-170">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="81d3f-170">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="81d3f-171">Experimente</span><span class="sxs-lookup"><span data-stu-id="81d3f-171">Try it out</span></span>

1. <span data-ttu-id="81d3f-172">Siga as instruções para a plataforma que você usará para executar o suplemento e fazer o sideload do suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="81d3f-172">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="81d3f-173">Windows: [Fazer o sideload de suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="81d3f-173">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="81d3f-174">Excel Online: [Fazer o sideload dos suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="81d3f-174">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="81d3f-175">iPad e Mac: [Fazer o sideload dos suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="81d3f-175">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="81d3f-176">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="81d3f-176">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="81d3f-178">Selecione qualquer intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="81d3f-178">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="81d3f-179">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="81d3f-179">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

---

## <a name="next-steps"></a><span data-ttu-id="81d3f-181">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="81d3f-181">Next steps</span></span>

<span data-ttu-id="81d3f-p116">Parabéns, você criou com sucesso um suplemento do Excel usando jQuery! Em seguida, aprenda mais sobre os recursos de um suplemento do Excel e crie um suplemento mais complexo seguindo o tutorial do suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="81d3f-p116">Congratulations, you've successfully created an Excel add-in using jQuery! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="81d3f-184">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="81d3f-184">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="81d3f-185">Confira também</span><span class="sxs-lookup"><span data-stu-id="81d3f-185">See also</span></span>

* [<span data-ttu-id="81d3f-186">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="81d3f-186">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="81d3f-187">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="81d3f-187">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="81d3f-188">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="81d3f-188">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="81d3f-189">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="81d3f-189">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
