---
title: Crie seu primeiro suplemento do painel de tarefas do Excel
description: Saiba como criar um Suplemento do Excel simples usando a API JS do Office.
ms.date: 07/17/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 781e2c3e7cd563e6ebeeaff3e8bf0624b64aec76
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/13/2019
ms.locfileid: "36308047"
---
# <a name="build-an-excel-task-pane-add-in"></a><span data-ttu-id="08433-103">Criar um suplemento do painel de tarefas do Excel</span><span class="sxs-lookup"><span data-stu-id="08433-103">Build an Excel task pane add-in using Angular</span></span>

<span data-ttu-id="08433-104">Neste artigo, você passará pelo processo de criação de um suplemento do painel de tarefas do Excel.</span><span class="sxs-lookup"><span data-stu-id="08433-104">In this article, you'll walk through the process of building an Outlook task pane add-in.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="08433-105">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="08433-105">Create the add-in</span></span>

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="yeoman-generatortabyeomangenerator"></a>[<span data-ttu-id="08433-106">Gerador do Yeoman</span><span class="sxs-lookup"><span data-stu-id="08433-106">Yeoman generator</span></span>](#tab/yeomangenerator)

### <a name="prerequisites"></a><span data-ttu-id="08433-107">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="08433-107">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="08433-108">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="08433-108">Create the add-in project</span></span>

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

<span data-ttu-id="08433-109">Use o gerador Yeoman para criar um projeto de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="08433-109">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="08433-110">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="08433-110">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="08433-111">**Escolha o tipo de projeto:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="08433-111">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="08433-112">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="08433-112">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="08433-113">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="08433-113">**What do you want to name your add-in?**</span></span> `my-office-add-in`
- <span data-ttu-id="08433-114">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="08433-114">**Which Office client application would you like to support?**</span></span> `Excel`

<span data-ttu-id="08433-115">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="08433-115">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

### <a name="explore-the-project"></a><span data-ttu-id="08433-116">Explore o projeto</span><span class="sxs-lookup"><span data-stu-id="08433-116">Explore the project</span></span>

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-js.md)]

### <a name="try-it-out"></a><span data-ttu-id="08433-117">Experimente</span><span class="sxs-lookup"><span data-stu-id="08433-117">Try it out</span></span>

1. <span data-ttu-id="08433-118">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="08433-118">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "my-office-add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="08433-119">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="08433-119">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="08433-121">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="08433-121">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="08433-122">Na parte inferior do painel de tarefas, escolha o link **Executar** para definir a cor do intervalo selecionado como amarelo.</span><span class="sxs-lookup"><span data-stu-id="08433-122">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-3c.png)

# <a name="visual-studiotabvisualstudio"></a>[<span data-ttu-id="08433-124">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="08433-124">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="08433-125">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="08433-125">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="08433-126">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="08433-126">Create the add-in project</span></span>

1. <span data-ttu-id="08433-127">Na barra de menus do Visual Studio, selecione **Arquivo**  >  **Novo**  >  **Projeto**.</span><span class="sxs-lookup"><span data-stu-id="08433-127">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="08433-128">Na lista de tipos de projeto em **Visual C#** ou em **Visual Basic**, expanda a opção **Office/SharePoint**, escolha **Suplementos** e depois **Suplemento da Web do Outlook** como o tipo de projeto.</span><span class="sxs-lookup"><span data-stu-id="08433-128">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="08433-129">Dê um nome ao projeto e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="08433-129">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="08433-130">Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel**e clique em **Concluir** para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="08433-130">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="08433-p102">O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="08433-p102">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="08433-133">Explorar a solução do Visual Studio</span><span class="sxs-lookup"><span data-stu-id="08433-133">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="08433-134">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="08433-134">Update the code</span></span>

1. <span data-ttu-id="08433-p103">**Home.html** especifica o HTML que será renderizado no painel de tarefas do suplemento. Em **Home.html**, substitua o elemento `<body>` pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="08433-p103">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>

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

2. <span data-ttu-id="08433-p104">Abra o arquivo **Home.js** na raiz do projeto do aplicativo Web. Este arquivo especifica o script do suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="08433-p104">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

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

3. <span data-ttu-id="08433-p105">Abra o arquivo **Home.css** na raiz do projeto do aplicativo Web. Este arquivo especifica os estilos personalizados para o suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="08433-p105">Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

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

### <a name="update-the-manifest"></a><span data-ttu-id="08433-143">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="08433-143">Update the manifest</span></span>

1. <span data-ttu-id="08433-144">Abra o arquivo de manifesto XML do projeto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="08433-144">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="08433-145">Este arquivo define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="08433-145">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="08433-p107">O elemento `ProviderName` tem um valor de espaço reservado. Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="08433-p107">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="08433-p108">O atributo `DefaultValue` do elemento `DisplayName` tem um espaço reservado. Substitua-o pelo **suplementos do My Office**.</span><span class="sxs-lookup"><span data-stu-id="08433-p108">The `DefaultValue` attribute of the `DisplayName` element has a placeholder. Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="08433-p109">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado. Substitua-o com **um suplemento do painel de tarefas do Excel**.</span><span class="sxs-lookup"><span data-stu-id="08433-p109">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="08433-152">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="08433-152">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="08433-153">Experimente</span><span class="sxs-lookup"><span data-stu-id="08433-153">Try it out</span></span>

1. <span data-ttu-id="08433-p110">Usando o Visual Studio, teste o suplemento do Excel recém-criado, pressionando **F5** ou escolhendo o botão **Iniciar** para abrir o Excel com o botão de suplemento **Mostrar Painel de Tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="08433-p110">Using Visual Studio, test the newly created Excel add-in by pressing **F5** or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="08433-156">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="08433-156">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="08433-158">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="08433-158">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="08433-159">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="08433-159">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

---

## <a name="next-steps"></a><span data-ttu-id="08433-161">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="08433-161">Next steps</span></span>

<span data-ttu-id="08433-162">Parabéns, você criou com êxito um suplemento do painel de tarefas do Excel!</span><span class="sxs-lookup"><span data-stu-id="08433-162">Congratulations, you've successfully created an Excel task pane add-in using Angular!</span></span> <span data-ttu-id="08433-163">Em seguida, saiba mais sobre os recursos de um suplemento do Excel e crie um suplemento mais complexo seguindo as etapas deste tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="08433-163">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="08433-164">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="08433-164">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="08433-165">Confira também</span><span class="sxs-lookup"><span data-stu-id="08433-165">See also</span></span>

* [<span data-ttu-id="08433-166">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="08433-166">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="08433-167">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="08433-167">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="08433-168">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="08433-168">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="08433-169">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="08433-169">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
