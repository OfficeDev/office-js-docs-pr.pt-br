---
title: Crie o seu primeiro suplemento do Excel usando JQuery
description: ''
ms.date: 01/17/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 7931e06899a94a0dcda2a5ab442d37ce21c119c0
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742440"
---
# <a name="build-an-excel-add-in-using-jquery"></a><span data-ttu-id="fd2a9-102">Criar um suplemento do Excel usando o jQuery</span><span class="sxs-lookup"><span data-stu-id="fd2a9-102">Build an Excel add-in using jQuery</span></span>

<span data-ttu-id="fd2a9-103">Neste artigo, você passará pelo processo de criar um suplemento do Excel usando o jQuery e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-103">In this article, you'll walk through the process of building an Excel add-in by using jQuery and the Excel JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="fd2a9-104">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="fd2a9-104">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="fd2a9-105">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="fd2a9-105">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="fd2a9-106">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="fd2a9-106">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="fd2a9-107">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="fd2a9-107">Create the add-in project</span></span>

1. <span data-ttu-id="fd2a9-108">Na barra de menus do Visual Studio, selecione **Arquivo**  >  **Novo**  >  **Projeto**.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="fd2a9-109">Na lista de tipos de projeto em **Visual C#** ou em **Visual Basic**, expanda a opção **Office/SharePoint**, escolha **Suplementos** e depois **Suplemento da Web do Outlook** como o tipo de projeto.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="fd2a9-110">Dê um nome ao projeto e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-110">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="fd2a9-111">Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel**e clique em **Concluir** para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-111">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="fd2a9-p101">O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="fd2a9-114">Explorar a solução do Visual Studio</span><span class="sxs-lookup"><span data-stu-id="fd2a9-114">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="fd2a9-115">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="fd2a9-115">Update the code</span></span>

1. <span data-ttu-id="fd2a9-116">**Home.html** especifica o HTML que será renderizado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-116">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="fd2a9-117">Em **Home.html**, substitua o elemento `<body>` pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-117">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
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

2. <span data-ttu-id="fd2a9-118">Abra o arquivo **Home.js** na raiz do projeto do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-118">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="fd2a9-119">Este arquivo especifica o script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-119">This file specifies the script for the add-in.</span></span> <span data-ttu-id="fd2a9-120">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-120">Replace the entire contents with the following code and save the file.</span></span> 

    ```js
    'use strict';

    (function () {

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                $('#set-color').click(setColor);
            });
        });

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

3. <span data-ttu-id="fd2a9-121">Abra o arquivo **Home.css** na raiz do projeto do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-121">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="fd2a9-122">Este arquivo especifica os estilos personalizados para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-122">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="fd2a9-123">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-123">Replace the entire contents with the following code and save the file.</span></span> 

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

### <a name="update-the-manifest"></a><span data-ttu-id="fd2a9-124">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="fd2a9-124">Update the manifest</span></span>

1. <span data-ttu-id="fd2a9-125">Abra o arquivo de manifesto XML do projeto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-125">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="fd2a9-126">Este arquivo define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-126">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="fd2a9-127">O elemento `ProviderName` tem um valor de espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-127">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="fd2a9-128">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-128">Replace it with your name.</span></span>

3. <span data-ttu-id="fd2a9-129">O atributo `DefaultValue` do elemento `DisplayName` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-129">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="fd2a9-130">Substitua-o pelo **suplementos do My Office**.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-130">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="fd2a9-131">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-131">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="fd2a9-132">Substitua-o com **um suplemento do painel de tarefas do Excel**.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-132">Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="fd2a9-133">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-133">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="fd2a9-134">Experimente</span><span class="sxs-lookup"><span data-stu-id="fd2a9-134">Try it out</span></span>

1. <span data-ttu-id="fd2a9-p109">Usando o Visual Studio, teste o suplemento do Excel recém-criado, pressionando **F5** ou escolhendo o botão **Iniciar** para abrir o Excel com o botão de suplemento **Mostrar Painel de Tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-p109">Using Visual Studio, test the newly created Excel add-in by pressing **F5** or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="fd2a9-137">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-137">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="fd2a9-139">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-139">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="fd2a9-140">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-140">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="fd2a9-142">Qualquer editor</span><span class="sxs-lookup"><span data-stu-id="fd2a9-142">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="fd2a9-143">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="fd2a9-143">Prerequisites</span></span>

- [<span data-ttu-id="fd2a9-144">Node.js</span><span class="sxs-lookup"><span data-stu-id="fd2a9-144">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="fd2a9-145">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-145">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a><span data-ttu-id="fd2a9-146">Criar o aplicativo Web</span><span class="sxs-lookup"><span data-stu-id="fd2a9-146">Create the web app</span></span>

1. <span data-ttu-id="fd2a9-147">Use o gerador Yeoman para criar um projeto de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-147">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="fd2a9-148">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="fd2a9-148">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="fd2a9-149">**Escolha o tipo de projeto:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="fd2a9-149">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="fd2a9-150">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="fd2a9-150">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="fd2a9-151">**Qual será o nome do suplemento?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="fd2a9-151">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="fd2a9-152">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="fd2a9-152">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Gerador do Yeoman](../images/yo-office-jquery.png)
    
    <span data-ttu-id="fd2a9-154">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-154">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="fd2a9-155">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-155">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

### <a name="update-the-code"></a><span data-ttu-id="fd2a9-156">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="fd2a9-156">Update the code</span></span> 

1. <span data-ttu-id="fd2a9-157">No editor de código, abra **index.html** na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-157">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="fd2a9-158">Esse arquivo especifica o HTML que será processado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-158">This file specifies the HTML that will be rendered in the add-in's task pane.</span></span> 
 
2. <span data-ttu-id="fd2a9-159">Dentro de **index.html**, substitua a marca `body` pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-159">Within **index.html**, replace the `body` tag with the following markup and save the file.</span></span>
 
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

3. <span data-ttu-id="fd2a9-160">Abra o arquivo **src\index.js** para especificar o script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-160">Open the file **src\index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="fd2a9-161">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-161">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';
    
    (function () {
        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                $('#set-color').click(setColor);
            });
        });

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

4. <span data-ttu-id="fd2a9-162">Abra o arquivo **app.css** para especificar os estilos personalizados do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-162">Open the file **app.css** to specify the custom styles for the add-in.</span></span> <span data-ttu-id="fd2a9-163">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-163">Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="fd2a9-164">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="fd2a9-164">Update the manifest</span></span>

1. <span data-ttu-id="fd2a9-165">Abra o arquivo **manifest.xml** para definir as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-165">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="fd2a9-166">O elemento `ProviderName` tem um valor de espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-166">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="fd2a9-167">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-167">Replace it with your name.</span></span>

3. <span data-ttu-id="fd2a9-168">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-168">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="fd2a9-169">Substitua-o com **um suplemento do painel de tarefas do Excel**.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-169">Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="fd2a9-170">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-170">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="fd2a9-171">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="fd2a9-171">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="fd2a9-172">Experimente</span><span class="sxs-lookup"><span data-stu-id="fd2a9-172">Try it out</span></span>

1. <span data-ttu-id="fd2a9-173">Siga as instruções para a plataforma que você usará para executar o suplemento e realizar sideload do suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-173">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="fd2a9-174">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="fd2a9-174">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="fd2a9-175">Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="fd2a9-175">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="fd2a9-176">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="fd2a9-176">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="fd2a9-177">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-177">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="fd2a9-179">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-179">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="fd2a9-180">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-180">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

---

## <a name="next-steps"></a><span data-ttu-id="fd2a9-182">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="fd2a9-182">Next steps</span></span>

<span data-ttu-id="fd2a9-p116">Você criou com êxito um suplemento do Excel usando jQuery!, parabéns! Agora, saiba mais sobre os recursos dos suplementos do Excel e crie um mais complexo, acompanhando o tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="fd2a9-p116">Congratulations, you've successfully created an Excel add-in using jQuery! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="fd2a9-185">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="fd2a9-185">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="fd2a9-186">Confira também</span><span class="sxs-lookup"><span data-stu-id="fd2a9-186">See also</span></span>

* [<span data-ttu-id="fd2a9-187">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="fd2a9-187">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="fd2a9-188">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="fd2a9-188">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="fd2a9-189">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="fd2a9-189">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="fd2a9-190">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="fd2a9-190">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)

