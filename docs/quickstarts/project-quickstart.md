---
title: Crie o seu primeiro suplemento do painel de tarefas do Project
description: ''
ms.date: 01/16/2020
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: cb2b76b989b62e3bb851f4a2b3ab27302fe0b21d
ms.sourcegitcommit: 8bce9c94540ed484d0749f07123dc7c72a6ca126
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/22/2020
ms.locfileid: "41265675"
---
# <a name="build-your-first-project-task-pane-add-in"></a><span data-ttu-id="b60ed-102">Crie o seu primeiro suplemento do painel de tarefas do Project</span><span class="sxs-lookup"><span data-stu-id="b60ed-102">Build your first Project task pane add-in</span></span>

<span data-ttu-id="b60ed-103">Neste artigo, você passará pelo processo de criação de um suplemento do painel de tarefas do Project.</span><span class="sxs-lookup"><span data-stu-id="b60ed-103">In this article, you'll walk through the process of building a Project task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="b60ed-104">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="b60ed-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="b60ed-105">Project 2016 ou posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="b60ed-105">Project 2016 or later on Windows</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="b60ed-106">Criar o suplemento</span><span class="sxs-lookup"><span data-stu-id="b60ed-106">Create the add-in</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="b60ed-107">**Escolha o tipo de projeto:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="b60ed-107">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="b60ed-108">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="b60ed-108">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="b60ed-109">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="b60ed-109">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="b60ed-110">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="b60ed-110">**Which Office client application would you like to support?**</span></span> `Project`

![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-project.png)

<span data-ttu-id="b60ed-112">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="b60ed-112">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="b60ed-113">Explore o projeto</span><span class="sxs-lookup"><span data-stu-id="b60ed-113">Explore the project</span></span>

<span data-ttu-id="b60ed-114">O projeto de suplemento que você criou com o gerador do Yeoman contém um exemplo de código para um suplemento de painel de tarefas bem básico.</span><span class="sxs-lookup"><span data-stu-id="b60ed-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="b60ed-115">O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="b60ed-115">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="b60ed-116">O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="b60ed-116">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="b60ed-117">O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="b60ed-117">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="b60ed-118">O arquivo **./src/taskpane/taskpane.js** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="b60ed-118">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="b60ed-119">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="b60ed-119">Update the code</span></span>

<span data-ttu-id="b60ed-120">No seu editor de código, abra o arquivo **./src/taskpane/taskpane.js** e adicione o seguinte código dentro da função **executar**.</span><span class="sxs-lookup"><span data-stu-id="b60ed-120">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="b60ed-121">Este código usa a API JavaScript do Office para configurar o `Name` campo e `Notes` campo da tarefa selecionada.</span><span class="sxs-lookup"><span data-stu-id="b60ed-121">This code uses the Office JavaScript API to set the `Name` field and `Notes` field of the selected task.</span></span>

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

## <a name="try-it-out"></a><span data-ttu-id="b60ed-122">Experimente</span><span class="sxs-lookup"><span data-stu-id="b60ed-122">Try it out</span></span>

1. <span data-ttu-id="b60ed-123">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="b60ed-123">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="b60ed-124">Inicie o servidor Web local.</span><span class="sxs-lookup"><span data-stu-id="b60ed-124">Start the local web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b60ed-125">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="b60ed-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="b60ed-126">Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="b60ed-126">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    <span data-ttu-id="b60ed-127">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="b60ed-127">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="b60ed-128">O servidor Web local é iniciado quando este comando é executado.</span><span class="sxs-lookup"><span data-stu-id="b60ed-128">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm start
    ```

3. <span data-ttu-id="b60ed-129">Em Project, crie um plano de projeto simples.</span><span class="sxs-lookup"><span data-stu-id="b60ed-129">In Project, create a simple project plan.</span></span>

4. <span data-ttu-id="b60ed-130">Carregue seu suplemento no Project seguindo as instruções em [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="b60ed-130">Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

5. <span data-ttu-id="b60ed-131">Selecione uma única tarefa dentro do projeto.</span><span class="sxs-lookup"><span data-stu-id="b60ed-131">Select a single task within the project.</span></span>

6. <span data-ttu-id="b60ed-132">Na parte inferior do painel de tarefas, escolha o link **Executar** para renomear a tarefa selecionada e adicionar anotações à tarefa selecionada.</span><span class="sxs-lookup"><span data-stu-id="b60ed-132">At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.</span></span>

    ![Captura de tela do aplicativo Project com o suplemento do painel de tarefas carregado](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a><span data-ttu-id="b60ed-134">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="b60ed-134">Next steps</span></span>

<span data-ttu-id="b60ed-135">Parabéns, você criou com êxito um suplemento do painel de tarefas do Project!</span><span class="sxs-lookup"><span data-stu-id="b60ed-135">Congratulations, you've successfully created a Project task pane add-in!</span></span> <span data-ttu-id="b60ed-136">Em seguida, saiba mais sobre os recursos de um suplemento do Project e explore os cenários comuns.</span><span class="sxs-lookup"><span data-stu-id="b60ed-136">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="b60ed-137">Suplementos do Project</span><span class="sxs-lookup"><span data-stu-id="b60ed-137">Project add-ins</span></span>](../project/project-add-ins.md)

## <a name="see-also"></a><span data-ttu-id="b60ed-138">Confira também</span><span class="sxs-lookup"><span data-stu-id="b60ed-138">See also</span></span>

- [<span data-ttu-id="b60ed-139">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="b60ed-139">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="b60ed-140">Principais conceitos dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b60ed-140">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="b60ed-141">Desenvolver Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="b60ed-141">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
