---
title: Crie o seu primeiro suplemento do painel de tarefas do Project
description: ''
ms.date: 05/08/2019
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: 466c8354a047c8e36741711881ad1b0e2993ddaf
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/13/2019
ms.locfileid: "36308005"
---
# <a name="build-your-first-project-task-pane-add-in"></a><span data-ttu-id="75484-102">Crie o seu primeiro suplemento do painel de tarefas do Project</span><span class="sxs-lookup"><span data-stu-id="75484-102">Build your first Project task pane add-in</span></span>

<span data-ttu-id="75484-103">Neste artigo, você passará pelo processo de criação de um suplemento do painel de tarefas do Project.</span><span class="sxs-lookup"><span data-stu-id="75484-103">In this article, you'll walk through the process of building a Project task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="75484-104">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="75484-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="75484-105">Project 2016 ou posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="75484-105">Project 2016 or later on Windows</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="75484-106">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="75484-106">Create the add-in</span></span>

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

<span data-ttu-id="75484-107">Use o gerador Yeoman para criar um projeto de suplemento do Project.</span><span class="sxs-lookup"><span data-stu-id="75484-107">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="75484-108">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="75484-108">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="75484-109">**Escolha o tipo de projeto:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="75484-109">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="75484-110">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="75484-110">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="75484-111">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="75484-111">**What do you want to name your add-in?**</span></span> `my-office-add-in`
- <span data-ttu-id="75484-112">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="75484-112">**Which Office client application would you like to support?**</span></span> `Project`

<span data-ttu-id="75484-113">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="75484-113">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

## <a name="explore-the-project"></a><span data-ttu-id="75484-114">Explore o projeto</span><span class="sxs-lookup"><span data-stu-id="75484-114">Explore the project</span></span>

<span data-ttu-id="75484-115">O projeto de suplemento que você criou com o gerador do Yeoman contém um exemplo de código para um suplemento de painel de tarefas bem básico.</span><span class="sxs-lookup"><span data-stu-id="75484-115">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="75484-116">O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="75484-116">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="75484-117">O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="75484-117">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="75484-118">O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="75484-118">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="75484-119">O arquivo **./src/taskpane/taskpane.js** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="75484-119">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="75484-120">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="75484-120">Update the code</span></span>

<span data-ttu-id="75484-121">No seu editor de código, abra o arquivo **./src/taskpane/taskpane.js** e adicione o seguinte código dentro da função **executar**.</span><span class="sxs-lookup"><span data-stu-id="75484-121">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="75484-122">Este código usa a API JavaScript do Office para configurar o `Name` campo e `Notes` campo da tarefa selecionada.</span><span class="sxs-lookup"><span data-stu-id="75484-122">This code uses the Office JavaScript API to set the `Name` field and `Notes` field of the selected task.</span></span>

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

## <a name="try-it-out"></a><span data-ttu-id="75484-123">Experimente</span><span class="sxs-lookup"><span data-stu-id="75484-123">Try it out</span></span>

1. <span data-ttu-id="75484-124">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="75484-124">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "my-office-add-in"
    ```

2. <span data-ttu-id="75484-125">Inicie o servidor Web local.</span><span class="sxs-lookup"><span data-stu-id="75484-125">Start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="75484-126">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="75484-126">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="75484-127">Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="75484-127">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    <span data-ttu-id="75484-128">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="75484-128">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="75484-129">O servidor Web local é iniciado quando este comando é executado.</span><span class="sxs-lookup"><span data-stu-id="75484-129">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm start
    ```

3. <span data-ttu-id="75484-130">Em Project, crie um plano de projeto simples.</span><span class="sxs-lookup"><span data-stu-id="75484-130">In Project, create a simple project plan.</span></span>

4. <span data-ttu-id="75484-131">Carregue seu suplemento no Project seguindo as instruções em [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="75484-131">Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

5. <span data-ttu-id="75484-132">Selecione uma única tarefa dentro do projeto.</span><span class="sxs-lookup"><span data-stu-id="75484-132">Select a single task within the project.</span></span>

6. <span data-ttu-id="75484-133">Na parte inferior do painel de tarefas, escolha o link **Executar** para renomear a tarefa selecionada e adicionar anotações à tarefa selecionada.</span><span class="sxs-lookup"><span data-stu-id="75484-133">At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.</span></span>

    ![Captura de tela do aplicativo Project com o suplemento do painel de tarefas carregado](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a><span data-ttu-id="75484-135">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="75484-135">Next steps</span></span>

<span data-ttu-id="75484-136">Parabéns, você criou com êxito um suplemento do painel de tarefas do Project!</span><span class="sxs-lookup"><span data-stu-id="75484-136">Congratulations, you've successfully created a Project task pane add-in!</span></span> <span data-ttu-id="75484-137">Em seguida, saiba mais sobre os recursos de um suplemento do Project e explore os cenários comuns.</span><span class="sxs-lookup"><span data-stu-id="75484-137">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="75484-138">Suplementos do Project</span><span class="sxs-lookup"><span data-stu-id="75484-138">Project add-ins</span></span>](../project/project-add-ins.md)

