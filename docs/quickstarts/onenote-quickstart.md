---
title: Crie seu primeiro suplemento do painel de tarefas do OneNote
description: ''
ms.date: 12/24/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 2e8c560aa02de690fa4e6abae25d0625379e26ad
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851562"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a><span data-ttu-id="ed8e3-102">Crie seu primeiro suplemento do painel de tarefas do OneNote</span><span class="sxs-lookup"><span data-stu-id="ed8e3-102">Build your first OneNote task pane add-in</span></span>

<span data-ttu-id="ed8e3-103">Neste artigo, você verá o processo de criação de um suplemento do painel de tarefas do OneNote.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-103">In this article, you'll walk through the process of building a OneNote task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ed8e3-104">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="ed8e3-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="ed8e3-105">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="ed8e3-105">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="ed8e3-106">**Escolha o tipo de projeto:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="ed8e3-106">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="ed8e3-107">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="ed8e3-107">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="ed8e3-108">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="ed8e3-108">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="ed8e3-109">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="ed8e3-109">**Which Office client application would you like to support?**</span></span> `OneNote`

![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-onenote.png)

<span data-ttu-id="ed8e3-111">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-111">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="ed8e3-112">Explore o projeto</span><span class="sxs-lookup"><span data-stu-id="ed8e3-112">Explore the project</span></span>

<span data-ttu-id="ed8e3-113">O projeto de suplemento que você criou com o gerador do Yeoman contém um exemplo de código para um suplemento de painel de tarefas bem básico.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-113">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="ed8e3-114">O arquivo **./manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-114">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="ed8e3-115">O arquivo **./src/taskpane/taskpane.html** contém a marcação HTML do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-115">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="ed8e3-116">O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-116">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="ed8e3-117">O arquivo **./src/taskpane/taskpane.js** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-117">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="ed8e3-118">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="ed8e3-118">Update the code</span></span>

<span data-ttu-id="ed8e3-119">No seu editor de código, abra o arquivo **./src/taskpane/taskpane.js** e adicione o seguinte código dentro da função **executar**.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-119">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="ed8e3-120">Este código usa a API JavaScript do OneNote para definir o título da página e adicionar um contorno ao corpo da página.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-120">This code uses the OneNote JavaScript API to set the page title and add an outline to the body of the page.</span></span>

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

## <a name="try-it-out"></a><span data-ttu-id="ed8e3-121">Experimente</span><span class="sxs-lookup"><span data-stu-id="ed8e3-121">Try it out</span></span>

1. <span data-ttu-id="ed8e3-122">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-122">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="ed8e3-123">Inicie o servidor Web local e realize o sideload no seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-123">Start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="ed8e3-124">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-124">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="ed8e3-125">Se você for solicitado a instalar um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-125">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="ed8e3-126">Se você estiver testando o seu suplemento no Mac, execute o seguinte comando antes de continuar.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-126">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="ed8e3-127">O servidor Web local é iniciado quando este comando é executado.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-127">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    <span data-ttu-id="ed8e3-128">Execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-128">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="ed8e3-129">Quando você executar este comando, o servidor da Web local será iniciado (se ainda não estiver em execução).</span><span class="sxs-lookup"><span data-stu-id="ed8e3-129">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

3. <span data-ttu-id="ed8e3-130">No [OneNote Online](https://www.onenote.com/notebooks), abra um bloco de anotações e crie uma nova página.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-130">In [OneNote on the web](https://www.onenote.com/notebooks), open a notebook and create a new page.</span></span>

4. <span data-ttu-id="ed8e3-131">Escolha **Inserir > Suplementos do Office** para abrir a caixa de diálogo Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-131">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="ed8e3-132">Se você estiver conectado à sua conta de consumidor, selecione a guia **MEUS SUPLEMENTOS** e escolha  **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-132">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="ed8e3-133">Se você estiver conectado à sua conta corporativa ou de estudante, selecione a guia **MINHA ORGANIZAÇÃO** e escolha  **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-133">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="ed8e3-134">A imagem a seguir mostra a guia **MEUS SUPLEMENTOS** para blocos de anotações do consumidor.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-134">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

5. <span data-ttu-id="ed8e3-135">Na caixa de diálogo Carregar Suplemento, navegue até **manifest.xml** na pasta do projeto e escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-135">In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

6. <span data-ttu-id="ed8e3-136">Na guia **Página Inicial**, na faixa de opções, escolha o botão **Mostrar Painel de Tarefas**.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-136">From the **Home** tab, choose the **Show Taskpane** button in the ribbon.</span></span> <span data-ttu-id="ed8e3-137">O painel de tarefa do suplemento abre em um iFrame ao lado da página do OneNote.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-137">The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

7. <span data-ttu-id="ed8e3-138">Na parte inferior do painel de tarefas, escolha o link **Executar** para definir o título da página e adicionar um contorno ao corpo da página.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-138">At the bottom of the task pane, choose the **Run** link to set the page title and add an outline to the body of the page.</span></span>

    ![O suplemento do OneNote criado a partir deste passo a passo](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a><span data-ttu-id="ed8e3-140">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="ed8e3-140">Next steps</span></span>

<span data-ttu-id="ed8e3-141">Parabéns, você criou com êxito um suplemento do painel de tarefas do OneNote!</span><span class="sxs-lookup"><span data-stu-id="ed8e3-141">Congratulations, you've successfully created a OneNote task pane add-in!</span></span> <span data-ttu-id="ed8e3-142">Em seguida, saiba mais sobre os principais conceitos de criação de suplementos do OneNote.</span><span class="sxs-lookup"><span data-stu-id="ed8e3-142">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="ed8e3-143">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="ed8e3-143">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="ed8e3-144">Confira também</span><span class="sxs-lookup"><span data-stu-id="ed8e3-144">See also</span></span>

* [<span data-ttu-id="ed8e3-145">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ed8e3-145">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="ed8e3-146">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="ed8e3-146">Building Office Add-ins using Office.js book</span></span>](../overview/office-add-ins-fundamentals.md)
* <span data-ttu-id="ed8e3-147">[Desenvolver Suplementos do Office ](../develop/develop-overview.md)</span><span class="sxs-lookup"><span data-stu-id="ed8e3-147">[](../develop/develop-overview.md)Develop Office Add-ins with Angular</span></span>
- [<span data-ttu-id="ed8e3-148">Visão geral da programação da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="ed8e3-148">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="ed8e3-149">Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="ed8e3-149">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="ed8e3-150">Amostra de Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="ed8e3-150">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)

