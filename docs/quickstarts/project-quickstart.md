---
title: Crie o seu primeiro suplemento do painel de tarefas do Project
description: ''
ms.date: 05/08/2019
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: ccc243b17b25dbdf4142e4a11086df78ef4a2670
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771734"
---
# <a name="build-your-first-project-task-pane-add-in"></a><span data-ttu-id="84c55-102">Crie o seu primeiro suplemento do painel de tarefas do Project</span><span class="sxs-lookup"><span data-stu-id="84c55-102">Build your first Project task pane add-in</span></span>

<span data-ttu-id="84c55-103">Neste artigo, você passará pelo processo de criação de um suplemento do painel de tarefas do Project.</span><span class="sxs-lookup"><span data-stu-id="84c55-103">In this article, you'll walk through the process of building a Project task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="84c55-104">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="84c55-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="84c55-105">Project 2016 ou posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="84c55-105">Project 2016 or later on Windows</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="84c55-106">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="84c55-106">Create the add-in</span></span>

<span data-ttu-id="84c55-107">Use o gerador Yeoman para criar um projeto de suplemento do Project.</span><span class="sxs-lookup"><span data-stu-id="84c55-107">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="84c55-108">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="84c55-108">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="84c55-109">**Escolha o tipo de projeto:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="84c55-109">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="84c55-110">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="84c55-110">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="84c55-111">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="84c55-111">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="84c55-112">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="84c55-112">**Which Office client application would you like to support?**</span></span> `Project`

![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-project.png)

<span data-ttu-id="84c55-114">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="84c55-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

## <a name="explore-the-project"></a><span data-ttu-id="84c55-115">Explore o projeto</span><span class="sxs-lookup"><span data-stu-id="84c55-115">Explore the project</span></span>

<span data-ttu-id="84c55-116">O projeto de suplemento que você criou com o gerador do Yeoman contém um exemplo de código para um suplemento de painel de tarefas bem básico.</span><span class="sxs-lookup"><span data-stu-id="84c55-116">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="84c55-117">O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="84c55-117">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="84c55-118">O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="84c55-118">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="84c55-119">O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="84c55-119">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="84c55-120">O arquivo **./src/taskpane/taskpane.js** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="84c55-120">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="84c55-121">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="84c55-121">Update the code</span></span>

<span data-ttu-id="84c55-122">No seu editor de código, abra o arquivo **./src/taskpane/taskpane.js** e adicione o seguinte código dentro da função **executar**.</span><span class="sxs-lookup"><span data-stu-id="84c55-122">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="84c55-123">Este código usa a API JavaScript do Office para configurar o `Name` campo e `Notes` campo da tarefa selecionada.</span><span class="sxs-lookup"><span data-stu-id="84c55-123">This code uses the Office JavaScript API to set the `Name` field and `Notes` field of the selected task.</span></span>

```js
var taskGuid;

// Get the GUID of the selected task
Office.context.document.getSelectedTaskAsync(
    function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            taskGuid = result.value;

            // Set the specified fields for the selected task.
            var targetFields = [Office.ProjectTaskFields.Name, Office.ProjectTaskFields.Notes];
            var fieldValues = ['New task name', 'Notes for the task.'];

            // Set the field value. If the call is successful, set the next field.
            for (var i = 0; i < targetFields.length; i++) {
                Office.context.document.setTaskFieldAsync(
                    taskGuid,
                    targetFields[i],
                    fieldValues[i],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            i++;
                        }
                        else {
                            var err = result.error;
                            console.log(err.name + ' ' + err.code + ' ' + err.message);
                        }
                    }
                );
            }
        } else {
            var err = result.error;
            console.log(err.name + ' ' + err.code + ' ' + err.message);
        }
    }
);
```

## <a name="try-it-out"></a><span data-ttu-id="84c55-124">Experimente</span><span class="sxs-lookup"><span data-stu-id="84c55-124">Try it out</span></span>

1. <span data-ttu-id="84c55-125">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="84c55-125">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="84c55-126">Inicie o servidor Web local.</span><span class="sxs-lookup"><span data-stu-id="84c55-126">Start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="84c55-127">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="84c55-127">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="84c55-128">Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="84c55-128">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    <span data-ttu-id="84c55-129">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="84c55-129">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="84c55-130">O servidor Web local é iniciado quando este comando é executado.</span><span class="sxs-lookup"><span data-stu-id="84c55-130">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm start
    ```

3. <span data-ttu-id="84c55-131">Em Project, crie um plano de projeto simples.</span><span class="sxs-lookup"><span data-stu-id="84c55-131">In Project, create a simple project plan.</span></span>

4. <span data-ttu-id="84c55-132">Carregue seu suplemento no Project seguindo as instruções em [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="84c55-132">Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

5. <span data-ttu-id="84c55-133">Selecione uma única tarefa dentro do projeto.</span><span class="sxs-lookup"><span data-stu-id="84c55-133">Select a single task within the project.</span></span>

6. <span data-ttu-id="84c55-134">Na parte inferior do painel de tarefas, escolha o link **Executar** para renomear a tarefa selecionada e adicionar anotações à tarefa selecionada.</span><span class="sxs-lookup"><span data-stu-id="84c55-134">At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.</span></span>

    ![Captura de tela do aplicativo Project com o suplemento do painel de tarefas carregado](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a><span data-ttu-id="84c55-136">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="84c55-136">Next steps</span></span>

<span data-ttu-id="84c55-137">Parabéns, você criou com êxito um suplemento do painel de tarefas do Project!</span><span class="sxs-lookup"><span data-stu-id="84c55-137">Congratulations, you've successfully created a Project task pane add-in!</span></span> <span data-ttu-id="84c55-138">Em seguida, saiba mais sobre os recursos de um suplemento do Project e explore os cenários comuns.</span><span class="sxs-lookup"><span data-stu-id="84c55-138">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="84c55-139">Suplementos do Project</span><span class="sxs-lookup"><span data-stu-id="84c55-139">Project add-ins</span></span>](../project/project-add-ins.md)

