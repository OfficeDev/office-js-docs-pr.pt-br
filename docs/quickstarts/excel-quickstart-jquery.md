---
title: Crie seu primeiro suplemento do painel de tarefas do Excel
description: Saiba como criar um Suplemento do Excel simples usando a API JS do Office.
ms.date: 04/03/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 01298244b11167b67a966dc6d8b66f9b3ba2a735
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185572"
---
# <a name="build-an-excel-task-pane-add-in"></a><span data-ttu-id="ceb52-103">Criar um suplemento do painel de tarefas do Excel</span><span class="sxs-lookup"><span data-stu-id="ceb52-103">Build an Excel task pane add-in</span></span>

<span data-ttu-id="ceb52-104">Neste artigo, você passará pelo processo de criação de um suplemento do painel de tarefas do Excel.</span><span class="sxs-lookup"><span data-stu-id="ceb52-104">In this article, you'll walk through the process of building an Excel task pane add-in.</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="ceb52-105">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="ceb52-105">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]
# <a name="yeoman-generator"></a>[<span data-ttu-id="ceb52-106">Gerador do Yeoman</span><span class="sxs-lookup"><span data-stu-id="ceb52-106">Yeoman generator</span></span>](#tab/yeomangenerator)

[!include[Redirect to the single sign-on (SSO) quick start](../includes/sso-quickstart-reference.md)]

## <a name="prerequisites"></a><span data-ttu-id="ceb52-107">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="ceb52-107">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="ceb52-108">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="ceb52-108">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="ceb52-109">**Escolha o tipo de projeto:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="ceb52-109">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="ceb52-110">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="ceb52-110">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="ceb52-111">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="ceb52-111">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="ceb52-112">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="ceb52-112">**Which Office client application would you like to support?**</span></span> `Excel`

![Gerador do Yeoman](../images/yo-office-excel.png)

<span data-ttu-id="ceb52-114">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="ceb52-114">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="explore-the-project"></a><span data-ttu-id="ceb52-115">Explore o projeto</span><span class="sxs-lookup"><span data-stu-id="ceb52-115">Explore the project</span></span>

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-js.md)]

### <a name="try-it-out"></a><span data-ttu-id="ceb52-116">Experimente</span><span class="sxs-lookup"><span data-stu-id="ceb52-116">Try it out</span></span>

1. <span data-ttu-id="ceb52-117">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="ceb52-117">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)]

3. <span data-ttu-id="ceb52-118">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="ceb52-118">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="ceb52-120">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="ceb52-120">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="ceb52-121">Na parte inferior do painel de tarefas, escolha o link **Executar** para definir a cor do intervalo selecionado como amarelo.</span><span class="sxs-lookup"><span data-stu-id="ceb52-121">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-3c.png)

### <a name="next-steps"></a><span data-ttu-id="ceb52-123">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="ceb52-123">Next steps</span></span>

<span data-ttu-id="ceb52-124">Parabéns, você criou com êxito um suplemento do painel de tarefas do Excel!</span><span class="sxs-lookup"><span data-stu-id="ceb52-124">Congratulations, you've successfully created an Excel task pane add-in!</span></span> <span data-ttu-id="ceb52-125">Em seguida, saiba mais sobre os recursos de um suplemento do Excel e crie um suplemento mais complexo seguindo as etapas deste [tutorial de suplemento do Excel](../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="ceb52-125">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the [Excel add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

# <a name="visual-studio"></a>[<span data-ttu-id="ceb52-126">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="ceb52-126">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="ceb52-127">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="ceb52-127">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="ceb52-128">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="ceb52-128">Create the add-in project</span></span>

1. <span data-ttu-id="ceb52-129">No Visual Studio, escolha **Criar um novo projeto**.</span><span class="sxs-lookup"><span data-stu-id="ceb52-129">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="ceb52-130">Usando a caixa de pesquisa, insira **suplemento**.</span><span class="sxs-lookup"><span data-stu-id="ceb52-130">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="ceb52-131">Escolha \*\*suplemento do Excel Web \*\*, em seguida, selecione **Próximo**.</span><span class="sxs-lookup"><span data-stu-id="ceb52-131">Choose **Excel Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="ceb52-132">Nomeie seu projeto e selecione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="ceb52-132">Name your project and select **Create**.</span></span>

4. <span data-ttu-id="ceb52-133">Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel**e clique em **Concluir** para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="ceb52-133">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="ceb52-p103">O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="ceb52-p103">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="ceb52-136">Explorar a solução do Visual Studio</span><span class="sxs-lookup"><span data-stu-id="ceb52-136">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="ceb52-137">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="ceb52-137">Update the code</span></span>

1. <span data-ttu-id="ceb52-p104">**Home.html** especifica o HTML que será renderizado no painel de tarefas do suplemento. Em **Home.html**, substitua o elemento `<body>` pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="ceb52-p104">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>

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

2. <span data-ttu-id="ceb52-p105">Abra o arquivo **Home.js** na raiz do projeto do aplicativo Web. Este arquivo especifica o script do suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="ceb52-p105">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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

3. <span data-ttu-id="ceb52-p106">Abra o arquivo **Home.css** na raiz do projeto do aplicativo Web. Este arquivo especifica os estilos personalizados para o suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="ceb52-p106">Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="ceb52-146">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="ceb52-146">Update the manifest</span></span>

1. <span data-ttu-id="ceb52-147">Abra o arquivo de manifesto XML do projeto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="ceb52-147">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="ceb52-148">Este arquivo define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="ceb52-148">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="ceb52-p108">O elemento `ProviderName` tem um valor de espaço reservado. Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="ceb52-p108">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="ceb52-p109">O atributo `DefaultValue` do elemento `DisplayName` tem um espaço reservado. Substitua-o pelo **suplementos do My Office**.</span><span class="sxs-lookup"><span data-stu-id="ceb52-p109">The `DefaultValue` attribute of the `DisplayName` element has a placeholder. Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="ceb52-p110">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado. Substitua-o com **um suplemento do painel de tarefas do Excel**.</span><span class="sxs-lookup"><span data-stu-id="ceb52-p110">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="ceb52-155">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="ceb52-155">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="ceb52-156">Experimente</span><span class="sxs-lookup"><span data-stu-id="ceb52-156">Try it out</span></span>

1. <span data-ttu-id="ceb52-157">Use o Visual Studio, teste o suplemento recém-criado do Excel pressionando **F5** ou escolha o botão **Iniciar** para iniciar o Excel com o botão do suplemento **Exibir painel de tarefas** exibido na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="ceb52-157">Using Visual Studio, test the newly created Excel add-in by pressing **F5** or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="ceb52-158">O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="ceb52-158">The add-in will be hosted locally on IIS.</span></span> <span data-ttu-id="ceb52-159">Se for solicitado que você confie em um certificado, faça-o para permitir que o suplemento se conecte ao seu organizador.</span><span class="sxs-lookup"><span data-stu-id="ceb52-159">If you are asked to trust a certificate, do so to allow the add-in to connect to its host.</span></span>

2. <span data-ttu-id="ceb52-160">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="ceb52-160">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="ceb52-162">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="ceb52-162">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="ceb52-163">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="ceb52-163">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

[!include[Console tool note](../includes/console-tool-note.md)]

### <a name="next-steps"></a><span data-ttu-id="ceb52-165">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="ceb52-165">Next steps</span></span>

<span data-ttu-id="ceb52-166">Parabéns, você criou com êxito um suplemento do painel de tarefas do Excel!</span><span class="sxs-lookup"><span data-stu-id="ceb52-166">Congratulations, you've successfully created an Excel task pane add-in!</span></span> <span data-ttu-id="ceb52-167">Em seguida, saiba mais sobre como [desenvolver suplementos do Office com o Visual Studio](../develop/develop-add-ins-visual-studio.md).</span><span class="sxs-lookup"><span data-stu-id="ceb52-167">Next, learn more about [developing Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

---

## <a name="see-also"></a><span data-ttu-id="ceb52-168">Confira também</span><span class="sxs-lookup"><span data-stu-id="ceb52-168">See also</span></span>

* [<span data-ttu-id="ceb52-169">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ceb52-169">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="ceb52-170">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="ceb52-170">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="ceb52-171">Desenvolver Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="ceb52-171">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="ceb52-172">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="ceb52-172">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="ceb52-173">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="ceb52-173">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="ceb52-174">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="ceb52-174">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
