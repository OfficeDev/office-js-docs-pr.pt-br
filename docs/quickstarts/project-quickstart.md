---
title: Criar o seu primeiro suplemento do Project
description: ''
ms.date: 01/17/2019
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: 4d0dfa98d36d6da56fe2b9687922371eea29062a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450760"
---
# <a name="build-your-first-project-add-in"></a><span data-ttu-id="c2a10-102">Criar o seu primeiro suplemento do Project</span><span class="sxs-lookup"><span data-stu-id="c2a10-102">Build your first Project add-in</span></span>

<span data-ttu-id="c2a10-103">Neste artigo, você passará pelo processo de criar um suplemento do Project usando o jQuery e a API JavaScript para Office.</span><span class="sxs-lookup"><span data-stu-id="c2a10-103">In this article, you'll walk through the process of building a Project add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c2a10-104">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="c2a10-104">Prerequisites</span></span>

- [<span data-ttu-id="c2a10-105">Node.js</span><span class="sxs-lookup"><span data-stu-id="c2a10-105">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="c2a10-106">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="c2a10-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a><span data-ttu-id="c2a10-107">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="c2a10-107">Create the add-in</span></span>

1. <span data-ttu-id="c2a10-108">Use o gerador Yeoman para criar um projeto de suplemento do Project.</span><span class="sxs-lookup"><span data-stu-id="c2a10-108">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="c2a10-109">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="c2a10-109">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="c2a10-110">**Escolha o tipo de projeto:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="c2a10-110">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="c2a10-111">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="c2a10-111">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="c2a10-112">**Qual será o nome do suplemento?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="c2a10-112">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="c2a10-113">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?:** `Project`</span><span class="sxs-lookup"><span data-stu-id="c2a10-113">**Which Office client application would you like to support?:** `Project`</span></span>

    ![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-project-jquery.png)
    
    <span data-ttu-id="c2a10-115">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="c2a10-115">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="c2a10-116">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="c2a10-116">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="c2a10-117">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="c2a10-117">Update the code</span></span>

1. <span data-ttu-id="c2a10-p102">No editor de código, abra **index.html** na raiz do projeto. Esse arquivo contém o HTML que será renderizado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c2a10-p102">In your code editor, open **index.html** in the root of the project. This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="c2a10-120">Substitua o elemento `<body>` pela seguinte marcação.</span><span class="sxs-lookup"><span data-stu-id="c2a10-120">Replace the `<body>` element with the following markup.</span></span>

    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Select a task and then choose the buttons below and observe the output in the <b>Results</b> textbox.</p>
                <h3>Try it out</h3>
                <button class="ms-Button" id="get-task-guid">Get Task GUID</button>
                <br/><br/>
                <button class="ms-Button" id="get-task">Get Task data</button>
                <br/>
                <h4>Results:</h4>
                <textarea id="result" rows="6" cols="25"></textarea>
            </div>
        </div>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. <span data-ttu-id="c2a10-121">Abra o arquivo **src/index.js** para especificar o script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c2a10-121">Open the file **src/index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="c2a10-122">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="c2a10-122">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        var taskGuid;

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                $('#get-task-guid').click(getTaskGUID);
                $('#get-task').click(getTask);
            });
        });

        function getTaskGUID() {
            Office.context.document.getSelectedTaskAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    result.value = "Task GUID: " + asyncResult.value;
                    taskGuid = asyncResult.value;
                }
                else {
                    console.log(asyncResult.error.message);
                }
            });
        }

        function getTask() {
            if (taskGuid != undefined) {
                Office.context.document.getTaskAsync(
                    taskGuid,
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            var taskInfo = asyncResult.value;
                            var taskOutput = "Task name: " + taskInfo.taskName +
                                            "\nGUID: " + taskGuid +
                                            "\nWSS Id: " + taskInfo.wssTaskId +
                                            "\nResource names: " + taskInfo.resourceNames;
                            result.value = taskOutput;
                        } else {
                            console.log(asyncResult.error.message);
                        }
                    }
                );
            } else {
                result.value = 'Task GUID not valid:\n' + taskGuid;
            } 
        }
    })();
    ```

4. <span data-ttu-id="c2a10-p104">Abra o arquivo **app.css** na raiz do projeto para especificar os estilos personalizados do suplemento. Substitua todo o conteúdo pelo que está a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="c2a10-p104">Open the file **app.css** in the root of the project to specify the custom styles for the add-in. Replace the entire contents with the following and save the file.</span></span>

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

## <a name="update-the-manifest"></a><span data-ttu-id="c2a10-125">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="c2a10-125">Update the manifest</span></span>

1. <span data-ttu-id="c2a10-126">Abra o arquivo **manifest.xml** para definir as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c2a10-126">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="c2a10-p105">O elemento `ProviderName` tem um valor de espaço reservado. Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="c2a10-p105">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="c2a10-p106">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado. Substitua-o com **um suplemento do painel de tarefas do Project**.</span><span class="sxs-lookup"><span data-stu-id="c2a10-p106">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Project**.</span></span>

4. <span data-ttu-id="c2a10-131">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="c2a10-131">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Project"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="c2a10-132">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="c2a10-132">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="c2a10-133">Experimente</span><span class="sxs-lookup"><span data-stu-id="c2a10-133">Try it out</span></span>

1. <span data-ttu-id="c2a10-134">No Project, crie um projeto simples que tenha pelo menos uma tarefa.</span><span class="sxs-lookup"><span data-stu-id="c2a10-134">In Project, create a simple project that has at least one task.</span></span>

2. <span data-ttu-id="c2a10-135">Siga as instruções para a plataforma que você usará para executar o suplemento e para fazer o sideload do suplemento no Project.</span><span class="sxs-lookup"><span data-stu-id="c2a10-135">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Project.</span></span>

    - <span data-ttu-id="c2a10-136">Windows: [Realizar o sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="c2a10-136">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="c2a10-137">Project Online: [Realizar o sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="c2a10-137">Project Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="c2a10-138">iPad e Mac: [Fazer sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="c2a10-138">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

3. <span data-ttu-id="c2a10-139">No Project, selecione uma tarefa.</span><span class="sxs-lookup"><span data-stu-id="c2a10-139">In Project, select a task.</span></span>

    ![Uma captura de tela de um plano de projeto no Project com uma tarefa selecionada](../images/project_quickstart_addin_1.png)

4. <span data-ttu-id="c2a10-141">No painel de tarefas, escolha o botão **Obter GUID de tarefas** para gravar a GUID de tarefas na caixa de texto **Resultados**.</span><span class="sxs-lookup"><span data-stu-id="c2a10-141">In the task pane, choose the **Get Task GUID** button to write the task GUID to the **Results** textbox.</span></span>

    ![Captura de tela de um plano de projeto no Project com uma tarefa selecionada e a GUID de tarefas gravada na caixa de texto no painel de tarefas](../images/project_quickstart_addin_2.png)

5. <span data-ttu-id="c2a10-143">No painel de tarefas, escolha o botão **Obter dados da tarefa** para gravar várias propriedades da tarefa selecionada na caixa de texto **Resultados**.</span><span class="sxs-lookup"><span data-stu-id="c2a10-143">In the task pane, choose the **Get Task data** button to write several properties of the selected task to the **Results** textbox.</span></span>

    ![Captura de tela de um plano de projeto no Project com uma tarefa selecionada e várias propriedades de tarefas gravadas na caixa de texto do painel de tarefas](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a><span data-ttu-id="c2a10-145">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="c2a10-145">Next steps</span></span>

<span data-ttu-id="c2a10-p107">Parabéns, você criou com êxito um suplemento do Project! Em seguida, saiba mais sobre os recursos de um suplemento do Project e explore os cenários comuns.</span><span class="sxs-lookup"><span data-stu-id="c2a10-p107">Congratulations, you've successfully created a Project add-in! Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="c2a10-148">Suplementos do Project</span><span class="sxs-lookup"><span data-stu-id="c2a10-148">Project add-ins</span></span>](../project/project-add-ins.md)

