---
title: Criar um suplemento do painel de tarefas do Excel usando o React
description: Aprenda a criar um suplemento do painel de tarefas simples do Excel usando a API do Office JS e reagir.
ms.date: 01/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: d98d3145a092b18bdfa60f29aa5dda885b9168ed
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596841"
---
# <a name="build-an-excel-task-pane-add-in-using-react"></a><span data-ttu-id="c2cb0-103">Criar um suplemento do painel de tarefas do Excel usando o React</span><span class="sxs-lookup"><span data-stu-id="c2cb0-103">Build an Excel task pane add-in using React</span></span>

<span data-ttu-id="c2cb0-104">Neste artigo, você passará pelo processo de criar um suplemento do painel de tarefas do Excel usando o React e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-104">In this article, you'll walk through the process of building an Excel task pane add-in using React and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c2cb0-105">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="c2cb0-105">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="c2cb0-106">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="c2cb0-106">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="c2cb0-107">**Escolha o tipo de projeto:** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="c2cb0-107">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="c2cb0-108">**Escolha o tipo de script:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="c2cb0-108">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="c2cb0-109">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="c2cb0-109">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="c2cb0-110">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="c2cb0-110">**Which Office client application would you like to support?**</span></span> `Excel`

![Gerador do Yeoman](../images/yo-office-excel-react-2.png)

<span data-ttu-id="c2cb0-112">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-112">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="c2cb0-113">Explore o projeto</span><span class="sxs-lookup"><span data-stu-id="c2cb0-113">Explore the project</span></span>

<span data-ttu-id="c2cb0-114">O projeto de suplemento que você criou com o gerador do Yeoman contém um exemplo de código para um suplemento de painel de tarefas bem básico.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="c2cb0-115">Se você quiser examinar os principais componentes do seu projeto de suplemento, abra o projeto no seu editor de código e revise os arquivos listados abaixo.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-115">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="c2cb0-116">Quando estiver pronto para experimentar o suplemento, prossiga para a próxima seção.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-116">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="c2cb0-117">O arquivo **manifest.xml** no diretório raiz do projeto define as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-117">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="c2cb0-118">O arquivo **./src/taskpane/taskpane.html** define a estrutura HTML do painel de tarefas e os arquivos na pasta **./src/taskpane/components** definem as diversas partes da interface do usuário do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-118">The **./src/taskpane/taskpane.html** file defines the HTML framework of the task pane, and the files within the **./src/taskpane/components** folder define the various parts of the task pane UI.</span></span>
- <span data-ttu-id="c2cb0-119">O arquivo **./src/taskpane/taskpane.css** contém o CSS que é aplicado ao conteúdo no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-119">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="c2cb0-120">O arquivo **./src/taskpane/components/App.tsx** contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o Excel.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-120">The **./src/taskpane/components/App.tsx** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="c2cb0-121">Experimente</span><span class="sxs-lookup"><span data-stu-id="c2cb0-121">Try it out</span></span>

1. <span data-ttu-id="c2cb0-122">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-122">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="c2cb0-123">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-123">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="c2cb0-125">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-125">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="c2cb0-126">Na parte inferior do painel de tarefas, escolha o link **Executar** para definir a cor do intervalo selecionado como amarelo.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-126">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="c2cb0-128">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="c2cb0-128">Next steps</span></span>

<span data-ttu-id="c2cb0-129">Parabéns, você criou com êxito um suplemento do painel de tarefas do Excel usando o React!</span><span class="sxs-lookup"><span data-stu-id="c2cb0-129">Congratulations, you've successfully created an Excel task pane add-in using React!</span></span> <span data-ttu-id="c2cb0-130">Em seguida, saiba mais sobre os recursos de um suplemento do Excel e crie um suplemento mais complexo seguindo as etapas deste tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="c2cb0-130">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="c2cb0-131">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="c2cb0-131">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="c2cb0-132">Confira também</span><span class="sxs-lookup"><span data-stu-id="c2cb0-132">See also</span></span>

* [<span data-ttu-id="c2cb0-133">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="c2cb0-133">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="c2cb0-134">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="c2cb0-134">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="c2cb0-135">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="c2cb0-135">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="c2cb0-136">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="c2cb0-136">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
