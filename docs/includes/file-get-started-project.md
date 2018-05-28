# <a name="build-your-first-project-add-in"></a><span data-ttu-id="d33ca-101">Criar o primeiro suplemento do Project</span><span class="sxs-lookup"><span data-stu-id="d33ca-101">Build your first Project add-in</span></span>

<span data-ttu-id="d33ca-102">Neste artigo, voc? passar? pelo processo de criar um suplemento do Project usando o jQuery e a API JavaScript para Office.</span><span class="sxs-lookup"><span data-stu-id="d33ca-102">In this article, you'll walk through the process of building a Project add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="d33ca-103">Pr?-requisitos</span><span class="sxs-lookup"><span data-stu-id="d33ca-103">Prerequisites</span></span>

- [<span data-ttu-id="d33ca-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="d33ca-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="d33ca-105">Instale a ?ltima vers?o do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="d33ca-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a><span data-ttu-id="d33ca-106">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="d33ca-106">Create the add-in</span></span>

1. <span data-ttu-id="d33ca-107">Crie uma pasta na sua unidade local e nomeie-a como `my-project-addin`.</span><span class="sxs-lookup"><span data-stu-id="d33ca-107">Create a folder on your local drive and name it `my-project-addin`.</span></span> <span data-ttu-id="d33ca-108">Esse ? o local em que voc? criar? os arquivos para seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="d33ca-108">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="d33ca-109">Navegue at? a nova pasta.</span><span class="sxs-lookup"><span data-stu-id="d33ca-109">Navigate to your new folder.</span></span>

    ```bash
    cd my-project-addin
    ```

3. <span data-ttu-id="d33ca-110">Use o gerador Yeoman para criar um projeto de suplemento do Project.</span><span class="sxs-lookup"><span data-stu-id="d33ca-110">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="d33ca-111">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="d33ca-111">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="d33ca-112">**Gostaria de criar uma nova subpasta para o seu projeto?** `No`</span><span class="sxs-lookup"><span data-stu-id="d33ca-112">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="d33ca-113">**Como deseja nomear seu suplemento?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="d33ca-113">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="d33ca-114">**Para qual aplicativo cliente do Office voc? deseja suporte?** `Project`</span><span class="sxs-lookup"><span data-stu-id="d33ca-114">**Which Office client application would you like to support?:** `Project`</span></span>
    - <span data-ttu-id="d33ca-115">**Gostaria de criar um novo suplemento?:** `Yes`</span><span class="sxs-lookup"><span data-stu-id="d33ca-115">**Would you like to create a new add-in?:** `Yes`</span></span>
    - <span data-ttu-id="d33ca-116">**Gostaria de usar o TypeScript?** `No`</span><span class="sxs-lookup"><span data-stu-id="d33ca-116">**Would you like to use TypeScript?:** `No`</span></span>
    - <span data-ttu-id="d33ca-117">**Escolha a estrutura:** `Jquery`</span><span class="sxs-lookup"><span data-stu-id="d33ca-117">**Choose a framework:** `Jquery`</span></span>

    <span data-ttu-id="d33ca-p103">O gerador perguntar? se voc? deseja abrir **resource.html**. N?o ? necess?rio abri-lo para este tutorial, mas fique ? vontade em fazer isso se tiver curiosidade. Escolha Sim ou N?o para concluir o assistente e deixar o gerador fazer seu trabalho.</span><span class="sxs-lookup"><span data-stu-id="d33ca-p103">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-project-jquery.png)

## <a name="update-the-code"></a><span data-ttu-id="d33ca-122">Atualizar o c?digo</span><span class="sxs-lookup"><span data-stu-id="d33ca-122">Update the code</span></span>

1. <span data-ttu-id="d33ca-123">No editor de c?digo, abra **index.html** na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="d33ca-123">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="d33ca-124">Esse arquivo cont?m o HTML que ser? renderizado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="d33ca-124">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="d33ca-125">Substitua o elemento `<header>` dentro do elemento `<body>` com a marca??o a seguir.</span><span class="sxs-lookup"><span data-stu-id="d33ca-125">Replace the `<header>` element inside the `<body>` element with the following markup.</span></span>

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

3. <span data-ttu-id="d33ca-126">Substitua o elemento `<main>` dentro do elemento `<body>` com a marca??o a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="d33ca-126">Replace the `<main>` element inside the `<body>` element with the following markup and save the file.</span></span>

    ```html
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
    ```

4. <span data-ttu-id="d33ca-127">Abra o arquivo **app.js** para especificar o script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="d33ca-127">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="d33ca-128">Substitua todo o conte?do pelo c?digo a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="d33ca-128">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        var taskGuid;

        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#get-task-guid').click(getTaskGUID);
                $('#get-task').click(getTask);
            });
        };

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

4. <span data-ttu-id="d33ca-129">Abra o arquivo **app.css** na raiz do projeto para especificar os estilos personalizados do suplemento.</span><span class="sxs-lookup"><span data-stu-id="d33ca-129">Open the file **app.css** in the root of the project to specify the custom styles for the add-in.</span></span> <span data-ttu-id="d33ca-130">Substitua todo o conte?do pelo que est? a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="d33ca-130">Replace the entire contents with the following and save the file.</span></span>

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

## <a name="update-the-manifest"></a><span data-ttu-id="d33ca-131">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="d33ca-131">Update the manifest</span></span>

1. <span data-ttu-id="d33ca-132">Abra o arquivo **my-office-add-in-manifest.xml** para definir as configura??es e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="d33ca-132">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="d33ca-133">O elemento `ProviderName` tem um valor de espa?o reservado.</span><span class="sxs-lookup"><span data-stu-id="d33ca-133">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="d33ca-134">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="d33ca-134">Replace it with your name.</span></span>

3. <span data-ttu-id="d33ca-135">O atributo `DefaultValue` do elemento `Description` tem um espa?o reservado.</span><span class="sxs-lookup"><span data-stu-id="d33ca-135">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="d33ca-136">Substitua-o com **um suplemento do painel de tarefas do Project**.</span><span class="sxs-lookup"><span data-stu-id="d33ca-136">Replace it with **A task pane add-in for Project**.</span></span>

4. <span data-ttu-id="d33ca-137">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="d33ca-137">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Project"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="d33ca-138">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="d33ca-138">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="d33ca-139">Experimente</span><span class="sxs-lookup"><span data-stu-id="d33ca-139">Try it out</span></span>

1. <span data-ttu-id="d33ca-140">No Project, crie um projeto simples que tenha pelo menos uma tarefa.</span><span class="sxs-lookup"><span data-stu-id="d33ca-140">In Project, create a simple project that has at least one task.</span></span>

2. <span data-ttu-id="d33ca-141">Siga as instru??es para a plataforma que voc? usar? para executar o suplemento e para fazer o sideload do suplemento no Project.</span><span class="sxs-lookup"><span data-stu-id="d33ca-141">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Project.</span></span>

    - <span data-ttu-id="d33ca-142">Windows: [Realizar o sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="d33ca-142">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="d33ca-143">Project Online: [Realizar o sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="d33ca-143">Project Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="d33ca-144">iPad e Mac: [Fazer sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="d33ca-144">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

3. <span data-ttu-id="d33ca-145">No Project, selecione uma tarefa.</span><span class="sxs-lookup"><span data-stu-id="d33ca-145">In Project, select a task.</span></span>

    ![Uma captura de tela de um plano de projeto no Project com uma tarefa selecionada](../images/project_quickstart_addin_1.png)

4. <span data-ttu-id="d33ca-147">No painel de tarefas, escolha o bot?o **Obter GUID de tarefas** para gravar a GUID de tarefas na caixa de texto **Resultados**.</span><span class="sxs-lookup"><span data-stu-id="d33ca-147">In the task pane, choose the **Get Task GUID** button to write the task GUID to the **Results** textbox.</span></span>

    ![Captura de tela de um plano de projeto no Project com uma tarefa selecionada e a GUID de tarefas gravada na caixa de texto no painel de tarefas](../images/project_quickstart_addin_2.png)

5. <span data-ttu-id="d33ca-149">No painel de tarefas, escolha o bot?o **Obter dados da tarefa** para gravar v?rias propriedades da tarefa selecionada na caixa de texto **Resultados**.</span><span class="sxs-lookup"><span data-stu-id="d33ca-149">In the task pane, choose the **Get Task data** button to write several properties of the selected task to the **Results** textbox.</span></span>

    ![Captura de tela de um plano de projeto no Project com uma tarefa selecionada e v?rias propriedades de tarefas gravadas na caixa de texto do painel de tarefas](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a><span data-ttu-id="d33ca-151">Pr?ximas etapas</span><span class="sxs-lookup"><span data-stu-id="d33ca-151">Next steps</span></span>

<span data-ttu-id="d33ca-152">Parab?ns, voc? criou com ?xito um suplemento do Project!</span><span class="sxs-lookup"><span data-stu-id="d33ca-152">Congratulations, you've successfully created a Project add-in!</span></span> <span data-ttu-id="d33ca-153">Em seguida, saiba mais sobre os recursos de um suplemento do Project e explore os cen?rios comuns.</span><span class="sxs-lookup"><span data-stu-id="d33ca-153">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="d33ca-154">Suplementos do Project</span><span class="sxs-lookup"><span data-stu-id="d33ca-154">Project add-ins</span></span>](../project/project-add-ins.md)
