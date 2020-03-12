---
title: Crie seu primeiro suplemento do painel de tarefas do Excel
description: Saiba como criar um Suplemento do Excel simples usando a API JS do Office.
ms.date: 01/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: cdd4de9cad88c09ec33e2cb1566b0a64afdf7745
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596617"
---
# <a name="build-an-excel-task-pane-add-in"></a><span data-ttu-id="4fd6b-103">Criar um suplemento do painel de tarefas do Excel</span><span class="sxs-lookup"><span data-stu-id="4fd6b-103">Build an Excel task pane add-in</span></span>

<span data-ttu-id="4fd6b-104">Neste artigo, você passará pelo processo de criação de um suplemento do painel de tarefas do Excel.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-104">In this article, you'll walk through the process of building an Excel task pane add-in.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="4fd6b-105">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="4fd6b-105">Create the add-in</span></span>

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="yeoman-generator"></a>[<span data-ttu-id="4fd6b-106">Gerador do Yeoman</span><span class="sxs-lookup"><span data-stu-id="4fd6b-106">Yeoman generator</span></span>](#tab/yeomangenerator)

[!include[Redirect to the single sign-on (SSO) quick start](../includes/sso-quickstart-reference.md)]

### <a name="prerequisites"></a><span data-ttu-id="4fd6b-107">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="4fd6b-107">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="4fd6b-108">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="4fd6b-108">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="4fd6b-109">**Escolha o tipo de projeto:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="4fd6b-109">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="4fd6b-110">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="4fd6b-110">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="4fd6b-111">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="4fd6b-111">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="4fd6b-112">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="4fd6b-112">**Which Office client application would you like to support?**</span></span> `Excel`

![Gerador do Yeoman](../images/yo-office-excel.png)

<span data-ttu-id="4fd6b-114">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-114">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="explore-the-project"></a><span data-ttu-id="4fd6b-115">Explore o projeto</span><span class="sxs-lookup"><span data-stu-id="4fd6b-115">Explore the project</span></span>

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-js.md)]

### <a name="try-it-out"></a><span data-ttu-id="4fd6b-116">Experimente</span><span class="sxs-lookup"><span data-stu-id="4fd6b-116">Try it out</span></span>

1. <span data-ttu-id="4fd6b-117">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-117">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="4fd6b-118">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-118">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="4fd6b-120">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-120">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="4fd6b-121">Na parte inferior do painel de tarefas, escolha o link **Executar** para definir a cor do intervalo selecionado como amarelo.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-121">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-3c.png)

### <a name="next-steps"></a><span data-ttu-id="4fd6b-123">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="4fd6b-123">Next steps</span></span>

<span data-ttu-id="4fd6b-124">Parabéns, você criou com êxito um suplemento do painel de tarefas do Excel!</span><span class="sxs-lookup"><span data-stu-id="4fd6b-124">Congratulations, you've successfully created an Excel task pane add-in!</span></span> <span data-ttu-id="4fd6b-125">Em seguida, saiba mais sobre os recursos de um suplemento do Excel e crie um suplemento mais complexo seguindo as etapas deste [tutorial de suplemento do Excel](../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="4fd6b-125">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the [Excel add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

# <a name="visual-studio"></a>[<span data-ttu-id="4fd6b-126">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="4fd6b-126">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="4fd6b-127">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="4fd6b-127">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="4fd6b-128">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="4fd6b-128">Create the add-in project</span></span>

1. <span data-ttu-id="4fd6b-129">No Visual Studio, escolha **Criar um novo projeto**.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-129">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="4fd6b-130">Usando a caixa de pesquisa, insira **suplemento**.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-130">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="4fd6b-131">Escolha \*\*suplemento do Excel Web \*\*, em seguida, selecione **Próximo**.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-131">Choose **Excel Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="4fd6b-132">Nomeie seu projeto e selecione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-132">Name your project and select **Create**.</span></span>

4. <span data-ttu-id="4fd6b-133">Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel**e clique em **Concluir** para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-133">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="4fd6b-p103">O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-p103">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="4fd6b-136">Explorar a solução do Visual Studio</span><span class="sxs-lookup"><span data-stu-id="4fd6b-136">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="4fd6b-137">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="4fd6b-137">Update the code</span></span>

1. <span data-ttu-id="4fd6b-p104">**Home.html** especifica o HTML que será renderizado no painel de tarefas do suplemento. Em **Home.html**, substitua o elemento `<body>` pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-p104">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>

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

2. <span data-ttu-id="4fd6b-p105">Abra o arquivo **Home.js** na raiz do projeto do aplicativo Web. Este arquivo especifica o script do suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-p105">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

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

3. <span data-ttu-id="4fd6b-p106">Abra o arquivo **Home.css** na raiz do projeto do aplicativo Web. Este arquivo especifica os estilos personalizados para o suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-p106">Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

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

### <a name="update-the-manifest"></a><span data-ttu-id="4fd6b-146">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="4fd6b-146">Update the manifest</span></span>

1. <span data-ttu-id="4fd6b-147">Abra o arquivo de manifesto XML do projeto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-147">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="4fd6b-148">Este arquivo define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-148">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="4fd6b-p108">O elemento `ProviderName` tem um valor de espaço reservado. Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-p108">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="4fd6b-p109">O atributo `DefaultValue` do elemento `DisplayName` tem um espaço reservado. Substitua-o pelo **suplementos do My Office**.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-p109">The `DefaultValue` attribute of the `DisplayName` element has a placeholder. Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="4fd6b-p110">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado. Substitua-o com **um suplemento do painel de tarefas do Excel**.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-p110">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="4fd6b-155">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-155">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="4fd6b-156">Experimente</span><span class="sxs-lookup"><span data-stu-id="4fd6b-156">Try it out</span></span>

1. <span data-ttu-id="4fd6b-p111">Usando o Visual Studio, teste o suplemento do Excel recém-criado, pressionando **F5** ou escolhendo o botão **Iniciar** para abrir o Excel com o botão de suplemento **Mostrar Painel de Tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-p111">Using Visual Studio, test the newly created Excel add-in by pressing **F5** or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="4fd6b-159">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-159">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="4fd6b-161">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-161">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="4fd6b-162">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="4fd6b-162">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

### <a name="next-steps"></a><span data-ttu-id="4fd6b-164">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="4fd6b-164">Next steps</span></span>

<span data-ttu-id="4fd6b-165">Parabéns, você criou com êxito um suplemento do painel de tarefas do Excel!</span><span class="sxs-lookup"><span data-stu-id="4fd6b-165">Congratulations, you've successfully created an Excel task pane add-in!</span></span> <span data-ttu-id="4fd6b-166">Em seguida, saiba mais sobre como [desenvolver suplementos do Office com o Visual Studio](../develop/develop-add-ins-visual-studio.md).</span><span class="sxs-lookup"><span data-stu-id="4fd6b-166">Next, learn more about [developing Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

---

## <a name="see-also"></a><span data-ttu-id="4fd6b-167">Confira também</span><span class="sxs-lookup"><span data-stu-id="4fd6b-167">See also</span></span>

* [<span data-ttu-id="4fd6b-168">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="4fd6b-168">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="4fd6b-169">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="4fd6b-169">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="4fd6b-170">Desenvolver Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="4fd6b-170">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="4fd6b-171">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="4fd6b-171">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="4fd6b-172">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="4fd6b-172">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="4fd6b-173">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="4fd6b-173">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
