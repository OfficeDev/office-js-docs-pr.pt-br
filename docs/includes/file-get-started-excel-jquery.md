# <a name="build-an-excel-add-in-using-jquery"></a><span data-ttu-id="5118c-101">Criar um suplemento do Excel usando o jQuery</span><span class="sxs-lookup"><span data-stu-id="5118c-101">Build an Excel add-in using jQuery</span></span>

<span data-ttu-id="5118c-102">Neste artigo, você passará pelo processo de criar um suplemento do Excel usando o jQuery e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="5118c-102">In this article, you'll walk through the process of building an Excel add-in by using jQuery and the Excel JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="5118c-103">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="5118c-103">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="5118c-104">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="5118c-104">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="5118c-105">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="5118c-105">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="5118c-106">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="5118c-106">Create the add-in project</span></span>

1. <span data-ttu-id="5118c-107">Na barra de menus do Visual Studio, selecione **Arquivo**  >  **Novo**  >  **Projeto**.</span><span class="sxs-lookup"><span data-stu-id="5118c-107">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="5118c-108">Na lista de tipos de projeto em **Visual C#** ou em **Visual Basic**, expanda a opção **Office/SharePoint**, escolha **Suplementos** e depois **Suplemento da Web do Outlook** como o tipo de projeto.</span><span class="sxs-lookup"><span data-stu-id="5118c-108">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="5118c-109">Dê um nome ao projeto e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="5118c-109">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="5118c-110">Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel**e clique em **Concluir** para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="5118c-110">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="5118c-p101">O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="5118c-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="5118c-113">Explorar a solução do Visual Studio</span><span class="sxs-lookup"><span data-stu-id="5118c-113">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="5118c-114">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="5118c-114">Update the code</span></span>

1. <span data-ttu-id="5118c-115">**Home.html** especifica o HTML que será renderizado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5118c-115">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="5118c-116">Em **Home.html**, substitua o elemento `<body>` pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="5118c-116">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
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

2. <span data-ttu-id="5118c-117">Abra o arquivo **Home.js** na raiz do projeto do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="5118c-117">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="5118c-118">Este arquivo especifica o script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5118c-118">This file specifies the script for the add-in.</span></span> <span data-ttu-id="5118c-119">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="5118c-119">Replace the entire contents with the following code and save the file.</span></span> 

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

3. <span data-ttu-id="5118c-120">Abra o arquivo **Home.css** na raiz do projeto do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="5118c-120">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="5118c-121">Este arquivo especifica os estilos personalizados para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="5118c-121">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="5118c-122">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="5118c-122">Replace the entire contents with the following code and save the file.</span></span> 

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

### <a name="update-the-manifest"></a><span data-ttu-id="5118c-123">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="5118c-123">Update the manifest</span></span>

1. <span data-ttu-id="5118c-124">Abra o arquivo de manifesto XML do projeto do Suplemento.</span><span class="sxs-lookup"><span data-stu-id="5118c-124">Open the XML manifest file in the Add-in project.</span></span> <span data-ttu-id="5118c-125">Este arquivo define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5118c-125">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="5118c-126">O elemento `ProviderName` tem um valor de espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="5118c-126">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="5118c-127">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="5118c-127">Replace it with your name.</span></span>

3. <span data-ttu-id="5118c-128">O atributo `DefaultValue` do elemento `DisplayName` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="5118c-128">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="5118c-129">Substitua-o pelo **suplementos do My Office**.</span><span class="sxs-lookup"><span data-stu-id="5118c-129">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="5118c-130">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="5118c-130">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="5118c-131">Substitua-o com **um suplemento do painel de tarefas do Excel**.</span><span class="sxs-lookup"><span data-stu-id="5118c-131">Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="5118c-132">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="5118c-132">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="5118c-133">Experimente</span><span class="sxs-lookup"><span data-stu-id="5118c-133">Try it out</span></span>

1. <span data-ttu-id="5118c-p109">Usando o Visual Studio, teste o suplemento do Excel recém-criado pressionando F5 ou escolhendo o botão **Iniciar** para abrir o Excel com o botão de suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="5118c-p109">Using Visual Studio, test the newly created Excel add-in by pressing F5 or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="5118c-136">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5118c-136">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do Suplemento do Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="5118c-138">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="5118c-138">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="5118c-139">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="5118c-139">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="5118c-141">Qualquer editor</span><span class="sxs-lookup"><span data-stu-id="5118c-141">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="5118c-142">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="5118c-142">Prerequisites</span></span>

- [<span data-ttu-id="5118c-143">Node.js</span><span class="sxs-lookup"><span data-stu-id="5118c-143">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="5118c-144">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="5118c-144">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a><span data-ttu-id="5118c-145">Criar o aplicativo Web</span><span class="sxs-lookup"><span data-stu-id="5118c-145">Create the web app</span></span>

1. <span data-ttu-id="5118c-146">Crie uma pasta na sua unidade local e nomeie-a como **my-addin**.</span><span class="sxs-lookup"><span data-stu-id="5118c-146">Create a folder on your local drive and name it **my-addin**.</span></span> <span data-ttu-id="5118c-147">Esse é o local em que você criará os arquivos para seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="5118c-147">This is where you'll create the files for your app.</span></span>

2. <span data-ttu-id="5118c-148">Navegue até a pasta do seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="5118c-148">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

3. <span data-ttu-id="5118c-149">Use o gerador do Yeoman para gerar o arquivo de manifesto para o seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="5118c-149">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="5118c-150">Execute o comando a seguir e responda aos prompts, conforme mostrado na seguinte captura de tela:</span><span class="sxs-lookup"><span data-stu-id="5118c-150">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="5118c-151">**Escolha um tipo de projeto:** `Jquery`</span><span class="sxs-lookup"><span data-stu-id="5118c-151">**Choose a project type:** `Jquery`</span></span>
    - <span data-ttu-id="5118c-152">**Escolha um tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="5118c-152">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="5118c-153">**Qual será o nome do suplemento?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="5118c-153">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="5118c-154">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="5118c-154">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Gerador do Yeoman](../images/yo-office-jquery.png)
    
    <span data-ttu-id="5118c-156">Depois de concluir o assistente, o gerador criará o projeto e instalará os componentes do nó de suporte.</span><span class="sxs-lookup"><span data-stu-id="5118c-156">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

    

4. <span data-ttu-id="5118c-157">No editor de código, abra **index.html** na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="5118c-157">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="5118c-158">Esse arquivo especifica o HTML que será renderizado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5118c-158">This file specifies the HTML that will be rendered in the add-in's task pane.</span></span> 
 
5. <span data-ttu-id="5118c-159">Dentro de **index.html**, substitua a marca `header` gerada pela seguinte marcação.</span><span class="sxs-lookup"><span data-stu-id="5118c-159">Within **index.html**, replace the generated `header` tag with the following markup.</span></span>
 
    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

6. <span data-ttu-id="5118c-160">Dentro de **index.html**, substitua a marca `main` gerada pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="5118c-160">Within **index.html**, replace the generated `main` tag with the following markup, and save the file.</span></span>

    ```html
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button class="ms-Button" id="set-color">Set color</button>
        </div>
    </div>
    ```

7. <span data-ttu-id="5118c-p113">Abra o arquivo **src\index.js** para especificar o script do suplemento. Substitua todo o conteúdo pelo seguinte código e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="5118c-p113">Open the file **app.js** to specify the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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

8. <span data-ttu-id="5118c-163">Abra o arquivo **app.css** para especificar os estilos personalizados do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5118c-163">Open the file **app.css** to specify the custom styles for the add-in.</span></span> <span data-ttu-id="5118c-164">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="5118c-164">Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="5118c-165">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="5118c-165">Update the manifest</span></span>

1. <span data-ttu-id="5118c-166">Abra o arquivo **my-office-add-in-manifest.xml** para definir as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5118c-166">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="5118c-167">O elemento `ProviderName` tem um valor de espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="5118c-167">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="5118c-168">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="5118c-168">Replace it with your name.</span></span>

3. <span data-ttu-id="5118c-169">O atributo `DefaultValue` do elemento `DisplayName` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="5118c-169">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="5118c-170">Substitua-o pelo **suplementos do My Office**.</span><span class="sxs-lookup"><span data-stu-id="5118c-170">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="5118c-171">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="5118c-171">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="5118c-172">Substitua-o com **um suplemento do painel de tarefas do Excel**.</span><span class="sxs-lookup"><span data-stu-id="5118c-172">Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="5118c-173">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="5118c-173">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="5118c-174">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="5118c-174">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="5118c-175">Experimente</span><span class="sxs-lookup"><span data-stu-id="5118c-175">Try it out</span></span>

1. <span data-ttu-id="5118c-176">Siga as instruções para a plataforma que você usará para executar o suplemento e realizar sideload do suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="5118c-176">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="5118c-177">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="5118c-177">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="5118c-178">Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="5118c-178">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="5118c-179">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="5118c-179">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="5118c-180">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5118c-180">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do Suplemento do Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="5118c-182">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="5118c-182">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="5118c-183">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="5118c-183">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

---

## <a name="next-steps"></a><span data-ttu-id="5118c-185">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="5118c-185">Next steps</span></span>

<span data-ttu-id="5118c-p118">Você criou com êxito um suplemento do Excel usando jQuery!, parabéns! Agora, saiba mais sobre os recursos dos suplementos do Excel e crie um mais complexo, acompanhando o tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="5118c-p118">Congratulations, you've successfully created an Excel add-in using jQuery! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="5118c-188">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="5118c-188">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="5118c-189">Veja também</span><span class="sxs-lookup"><span data-stu-id="5118c-189">See also</span></span>

* [<span data-ttu-id="5118c-190">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="5118c-190">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="5118c-191">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="5118c-191">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="5118c-192">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="5118c-192">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="5118c-193">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="5118c-193">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
