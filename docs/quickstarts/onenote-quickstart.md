---
title: Crie seu primeiro suplemento do painel de tarefas do OneNote
description: ''
ms.date: 06/20/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 24c8854cb1f9332371f3726409f91f7cdbf53243
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/13/2019
ms.locfileid: "36308019"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a><span data-ttu-id="59e00-102">Crie seu primeiro suplemento do painel de tarefas do OneNote</span><span class="sxs-lookup"><span data-stu-id="59e00-102">Build your first Word task pane add-in</span></span>

<span data-ttu-id="59e00-103">Neste artigo, você verá o processo de criação de um suplemento do painel de tarefas do OneNote.</span><span class="sxs-lookup"><span data-stu-id="59e00-103">In this article, you'll walk through the process of building a Project task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="59e00-104">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="59e00-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="59e00-105">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="59e00-105">Create the add-in project</span></span>

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

<span data-ttu-id="59e00-106">Use o gerador Yeoman para criar um projeto de suplemento do OneNote.</span><span class="sxs-lookup"><span data-stu-id="59e00-106">Use the Yeoman generator to create a OneNote add-in project.</span></span> <span data-ttu-id="59e00-107">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="59e00-107">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="59e00-108">**Escolha o tipo de projeto:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="59e00-108">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="59e00-109">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="59e00-109">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="59e00-110">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="59e00-110">**What do you want to name your add-in?**</span></span> `my-office-add-in`
- <span data-ttu-id="59e00-111">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="59e00-111">**Which Office client application would you like to support?**</span></span> `OneNote`

<span data-ttu-id="59e00-112">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="59e00-112">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
## <a name="explore-the-project"></a><span data-ttu-id="59e00-113">Explore o projeto</span><span class="sxs-lookup"><span data-stu-id="59e00-113">Explore the project</span></span>

<span data-ttu-id="59e00-114">O projeto de suplemento que você criou com o gerador do Yeoman contém um exemplo de código para um suplemento de painel de tarefas bem básico.</span><span class="sxs-lookup"><span data-stu-id="59e00-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="59e00-115">O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="59e00-115">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="59e00-116">O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="59e00-116">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="59e00-117">O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="59e00-117">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="59e00-118">O arquivo **./src/taskpane/taskpane.js** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="59e00-118">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="59e00-119">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="59e00-119">Update the code</span></span>

<span data-ttu-id="59e00-120">No seu editor de código, abra o arquivo **./src/taskpane/taskpane.js** e adicione o seguinte código dentro da função **executar**.</span><span class="sxs-lookup"><span data-stu-id="59e00-120">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="59e00-121">Este código usa a API JavaScript do OneNote para definir o título da página e adicionar um contorno ao corpo da página.</span><span class="sxs-lookup"><span data-stu-id="59e00-121">This code uses the OneNote JavaScript API to set the page title and add an outline to the body of the page.</span></span>

```js
try {
    await OneNote.run(async context => {

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to set the page title.
        page.title = "Hello World";

        // Queue a command to add an outline to the page.
        var html = "<p><ol><li>Item #1</li><li>Item #2</li></ol></p>";
        page.addOutline(40, 90, html);

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync();
    });
} catch (error) {
    console.log("Error: " + error);
}
```

## <a name="try-it-out"></a><span data-ttu-id="59e00-122">Experimente</span><span class="sxs-lookup"><span data-stu-id="59e00-122">Try it out</span></span>

1. <span data-ttu-id="59e00-123">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="59e00-123">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "my-office-add-in"
    ```

2. <span data-ttu-id="59e00-124">Inicie o servidor Web local e realize o sideload no seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="59e00-124">Start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="59e00-125">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="59e00-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="59e00-126">Se você for solicitado a instalar um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="59e00-126">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="59e00-127">Se você estiver testando o seu suplemento no Mac, execute o seguinte comando antes de continuar.</span><span class="sxs-lookup"><span data-stu-id="59e00-127">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="59e00-128">O servidor Web local é iniciado quando este comando é executado.</span><span class="sxs-lookup"><span data-stu-id="59e00-128">When you run this command, the local web server will start.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    <span data-ttu-id="59e00-129">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="59e00-129">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="59e00-130">Quando você executar este comando, o servidor da Web local será iniciado (se ainda não estiver em execução).</span><span class="sxs-lookup"><span data-stu-id="59e00-130">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

3. <span data-ttu-id="59e00-131">No [OneNote Online](https://www.onenote.com/notebooks), abra um bloco de anotações e crie uma nova página.</span><span class="sxs-lookup"><span data-stu-id="59e00-131">In [OneNote on the web](https://www.onenote.com/notebooks), open a notebook and create a new page.</span></span>

4. <span data-ttu-id="59e00-132">Escolha **Inserir > Suplementos do Office** para abrir a caixa de diálogo Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="59e00-132">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="59e00-133">Se você estiver conectado à sua conta de consumidor, selecione a guia **MEUS SUPLEMENTOS** e escolha  **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="59e00-133">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="59e00-134">Se você estiver conectado à sua conta corporativa ou de estudante, selecione a guia **MINHA ORGANIZAÇÃO** e escolha  **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="59e00-134">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="59e00-135">A imagem a seguir mostra a guia **MEUS SUPLEMENTOS** para blocos de anotações do consumidor.</span><span class="sxs-lookup"><span data-stu-id="59e00-135">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

5. <span data-ttu-id="59e00-136">Na caixa de diálogo Carregar Suplemento, navegue até **manifest.xml** na pasta do projeto e escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="59e00-136">In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

6. <span data-ttu-id="59e00-137">Na guia **Página Inicial**, na faixa de opções, escolha o botão **Mostrar Painel de Tarefas**.</span><span class="sxs-lookup"><span data-stu-id="59e00-137">From the **Home** tab, choose the **Show Taskpane** button in the ribbon.</span></span> <span data-ttu-id="59e00-138">O painel de tarefa do suplemento abre em um iFrame ao lado da página do OneNote.</span><span class="sxs-lookup"><span data-stu-id="59e00-138">The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

7. <span data-ttu-id="59e00-139">Na parte inferior do painel de tarefas, escolha o link **Executar** para definir o título da página e adicionar um contorno ao corpo da página.</span><span class="sxs-lookup"><span data-stu-id="59e00-139">At the bottom of the task pane, choose the **Run** link to set the page title and add an outline to the body of the page.</span></span>

    ![O suplemento do OneNote criado a partir deste passo a passo](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a><span data-ttu-id="59e00-141">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="59e00-141">Next steps</span></span>

<span data-ttu-id="59e00-142">Parabéns, você criou com êxito um suplemento do painel de tarefas do OneNote!</span><span class="sxs-lookup"><span data-stu-id="59e00-142">Congratulations, you've successfully created a Word task pane add-in!</span></span> <span data-ttu-id="59e00-143">Em seguida, saiba mais sobre os principais conceitos de criação de suplementos do OneNote.</span><span class="sxs-lookup"><span data-stu-id="59e00-143">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="59e00-144">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="59e00-144">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="59e00-145">Confira também</span><span class="sxs-lookup"><span data-stu-id="59e00-145">See also</span></span>

- [<span data-ttu-id="59e00-146">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="59e00-146">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="59e00-147">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="59e00-147">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="59e00-148">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="59e00-148">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="59e00-149">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="59e00-149">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

