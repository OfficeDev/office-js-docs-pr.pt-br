---
title: Tutorial de suplemento do Word
description: Neste tutorial, voc? criar? um suplemento do Word que insere (e substitui) intervalos de texto, par?grafos, imagens, HTML, tabelas e controles de conte?do. Você também aprenderá como formatar texto e como inserir (e substituir) conteúdo nos controles de conteúdo.
ms.date: 07/07/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 9ee851c9d479c15a0abce5228d89648d1268861b
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093508"
---
# <a name="tutorial-create-a-word-task-pane-add-in"></a><span data-ttu-id="688a4-104">Tutorial: Criar Suplemento do Painel de Tarefas no Word</span><span class="sxs-lookup"><span data-stu-id="688a4-104">Tutorial: Create a Word task pane add-in</span></span>

<span data-ttu-id="688a4-105">Neste tutorial: você criará um suplemento do painel de tarefas no Word:</span><span class="sxs-lookup"><span data-stu-id="688a4-105">In this tutorial, you'll create a Word task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="688a4-106">Insere um intervalo de texto</span><span class="sxs-lookup"><span data-stu-id="688a4-106">Inserts a range of text</span></span>
> * <span data-ttu-id="688a4-107">Formatos de texto</span><span class="sxs-lookup"><span data-stu-id="688a4-107">Formats text</span></span>
> * <span data-ttu-id="688a4-108">Substitui e insere texto em vários locais</span><span class="sxs-lookup"><span data-stu-id="688a4-108">Replaces text and inserts text in various locations</span></span>
> * <span data-ttu-id="688a4-109">Insere imagens, HTML e tabelas</span><span class="sxs-lookup"><span data-stu-id="688a4-109">Inserts images, HTML, and tables</span></span>
> * <span data-ttu-id="688a4-110">Cria e atualiza os controles de conteúdo</span><span class="sxs-lookup"><span data-stu-id="688a4-110">Creates and updates content controls</span></span> 

> [!TIP]
> <span data-ttu-id="688a4-111">Se você já concluiu o início rápido [Criar um suplemento do painel de tarefas do Word](../quickstarts/word-quickstart.md) e deseja usar esse projeto como ponto de partida para este tutorial, vá diretamente para a seção [Inserir um intervalo de texto](#insert-a-range-of-text) para iniciar o tutorial.</span><span class="sxs-lookup"><span data-stu-id="688a4-111">If you've already completed the [Build your first Word task pane add-in](../quickstarts/word-quickstart.md) quick start, and want to use that project as a starting point for this tutorial, go directly to the [Insert a range of text](#insert-a-range-of-text) section to start this tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="688a4-112">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="688a4-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-your-add-in-project"></a><span data-ttu-id="688a4-113">Criar seu projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="688a4-113">Create your add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="688a4-114">**Escolha o tipo de projeto:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="688a4-114">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="688a4-115">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="688a4-115">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="688a4-116">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="688a4-116">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="688a4-117">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="688a4-117">**Which Office client application would you like to support?**</span></span> `Word`

![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-word.png)

<span data-ttu-id="688a4-119">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="688a4-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="insert-a-range-of-text"></a><span data-ttu-id="688a4-120">Inserir um intervalo de texto</span><span class="sxs-lookup"><span data-stu-id="688a4-120">Insert a range of text</span></span>

<span data-ttu-id="688a4-121">Nesta etapa do tutorial, você testará programaticamente se o suplemento oferece suporte à versão atual do Word do usuário e inserirá um parágrafo no documento.</span><span class="sxs-lookup"><span data-stu-id="688a4-121">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Word, and then insert a paragraph into the document.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="688a4-122">Codificação do suplemento</span><span class="sxs-lookup"><span data-stu-id="688a4-122">Code the add-in</span></span>

1. <span data-ttu-id="688a4-123">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="688a4-123">Open the project in your code editor.</span></span>

2. <span data-ttu-id="688a4-124">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="688a4-124">Open the file **./src/taskpane/taskpane.html**.</span></span> <span data-ttu-id="688a4-125">Ele contém a marcação HTML para o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="688a4-125">This file contains the HTML markup for the task pane.</span></span>

3. <span data-ttu-id="688a4-126">Localize o elemento `<main>` e exclua todas as linhas que aparecem após a marca de abertura `<main>` e antes da marca de fechamento `</main>`.</span><span class="sxs-lookup"><span data-stu-id="688a4-126">Locate the `<main>` element and delete all lines that appear after the opening `<main>` tag and before the closing `</main>` tag.</span></span>

4. <span data-ttu-id="688a4-127">Adicione a seguinte marcação imediatamente após a marca de abertura `<main>`:</span><span class="sxs-lookup"><span data-stu-id="688a4-127">Add the following markup immediately after the opening `<main>` tag:</span></span>

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button><br/><br/>
    ```

5. <span data-ttu-id="688a4-128">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-128">Open the file **./src/taskpane/taskpane.js**.</span></span> <span data-ttu-id="688a4-129">Ele contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="688a4-129">This file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

6. <span data-ttu-id="688a4-130">Remova todas as referências ao botão `run` e à função `run()` da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="688a4-130">Remove all references to the `run` button and the `run()` function by doing the following:</span></span>

    - <span data-ttu-id="688a4-131">Localize e exclua a linha `document.getElementById("run").onclick = run;`.</span><span class="sxs-lookup"><span data-stu-id="688a4-131">Locate and delete the line `document.getElementById("run").onclick = run;`.</span></span>

    - <span data-ttu-id="688a4-132">Localize e exclua toda a função `run()`.</span><span class="sxs-lookup"><span data-stu-id="688a4-132">Locate and delete the entire `run()` function.</span></span>

7. <span data-ttu-id="688a4-133">Na chamada do método `Office.onReady`, localize a linha `if (info.host === Office.HostType.Word) {` e adicione o seguinte código imediatamente após ela.</span><span class="sxs-lookup"><span data-stu-id="688a4-133">Within the `Office.onReady` method call, locate the line `if (info.host === Office.HostType.Word) {` and add the following code immediately after that line.</span></span> <span data-ttu-id="688a4-134">Observação:</span><span class="sxs-lookup"><span data-stu-id="688a4-134">Note:</span></span>

    - <span data-ttu-id="688a4-135">A primeira parte desse código determina se a versão do Word do usuário suporta uma versão do Word.js que inclui todas as APIs usadas em todos os estágios desse tutorial.</span><span class="sxs-lookup"><span data-stu-id="688a4-135">The first part of this code determines whether the user's version of Word supports a version of Word.js that includes all the APIs that are used in all stages of this tutorial.</span></span> <span data-ttu-id="688a4-136">Em um suplemento de produção, use o corpo do bloco condicional para ocultar ou desabilitar a interface do usuário que chame a APIs sem suporte.</span><span class="sxs-lookup"><span data-stu-id="688a4-136">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="688a4-137">Isso permitirá que o usuário ainda use as partes do suplemento às quais a versão do Word dá suporte.</span><span class="sxs-lookup"><span data-stu-id="688a4-137">This will enable the user to still use the parts of the add-in that are supported by their version of Word.</span></span>
    - <span data-ttu-id="688a4-138">A segunda parte desse código adiciona um manipulador de eventos para o botão `insert-paragraph`.</span><span class="sxs-lookup"><span data-stu-id="688a4-138">The second part of this code adds an event handler for the `insert-paragraph` button.</span></span>

    ```js
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = insertParagraph;
    ```

8. <span data-ttu-id="688a4-139">Adicione a seguinte função ao final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="688a4-139">Add the following function to the end of the file.</span></span> <span data-ttu-id="688a4-140">Observação:</span><span class="sxs-lookup"><span data-stu-id="688a4-140">Note:</span></span>

   - <span data-ttu-id="688a4-141">Your Word.js business logic will be added to the function that is passed to `Word.run`.</span><span class="sxs-lookup"><span data-stu-id="688a4-141">Your Word.js business logic will be added to the function that is passed to `Word.run`.</span></span> <span data-ttu-id="688a4-142">This logic does not execute immediately.</span><span class="sxs-lookup"><span data-stu-id="688a4-142">This logic does not execute immediately.</span></span> <span data-ttu-id="688a4-143">Instead, it is added to a queue of pending commands.</span><span class="sxs-lookup"><span data-stu-id="688a4-143">Instead, it is added to a queue of pending commands.</span></span>

   - <span data-ttu-id="688a4-144">O método `context.sync` envia todos os comandos da fila para execução no Word.</span><span class="sxs-lookup"><span data-stu-id="688a4-144">The `context.sync` method sends all queued commands to Word for execution.</span></span>

   - <span data-ttu-id="688a4-145">The `Word.run` is followed by a `catch` block.</span><span class="sxs-lookup"><span data-stu-id="688a4-145">The `Word.run` is followed by a `catch` block.</span></span> <span data-ttu-id="688a4-146">This is a best practice that you should always follow.</span><span class="sxs-lookup"><span data-stu-id="688a4-146">This is a best practice that you should always follow.</span></span> 

    ```js
    function insertParagraph() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert a paragraph into the document.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

9. <span data-ttu-id="688a4-147">Na função `insertParagraph()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-147">Within the `insertParagraph()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="688a4-148">Observação:</span><span class="sxs-lookup"><span data-stu-id="688a4-148">Note:</span></span>

   - <span data-ttu-id="688a4-149">O primeiro parâmetro para o método `insertParagraph` é o texto para o novo parágrafo.</span><span class="sxs-lookup"><span data-stu-id="688a4-149">The first parameter to the `insertParagraph` method is the text for the new paragraph.</span></span>

   - <span data-ttu-id="688a4-150">The second parameter is the location within the body where the paragraph will be inserted.</span><span class="sxs-lookup"><span data-stu-id="688a4-150">The second parameter is the location within the body where the paragraph will be inserted.</span></span> <span data-ttu-id="688a4-151">Other options for insert paragraph, when the parent object is the body, are "End" and "Replace".</span><span class="sxs-lookup"><span data-stu-id="688a4-151">Other options for insert paragraph, when the parent object is the body, are "End" and "Replace".</span></span>

    ```js
    var docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
                            "Start");
    ```

10. <span data-ttu-id="688a4-152">Verifique se você salvou todas as alterações feitas no projeto.</span><span class="sxs-lookup"><span data-stu-id="688a4-152">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="688a4-153">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="688a4-153">Test the add-in</span></span>

1. <span data-ttu-id="688a4-154">Conclua as etapas a seguir para iniciar o servidor Web local e fazer o sideload do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="688a4-154">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="688a4-155">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="688a4-155">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="688a4-156">Se você for solicitado a instalar um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="688a4-156">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="688a4-157">Se você estiver testando seu suplemento no Mac, execute o seguinte comando no diretório raiz do seu projeto antes de continuar.</span><span class="sxs-lookup"><span data-stu-id="688a4-157">If you're testing your add-in on Mac, run the following command in the root directory of your project before proceeding.</span></span> <span data-ttu-id="688a4-158">O servidor Web local é iniciado quando este comando é executado.</span><span class="sxs-lookup"><span data-stu-id="688a4-158">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="688a4-159">Para testar seu suplemento no Word, execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="688a4-159">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="688a4-160">Isso inicia o servidor Web local (caso ainda não esteja em execução) e abre o Word com o suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="688a4-160">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="688a4-161">Para testar o suplemento no Word na Web, execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="688a4-161">To test your add-in in Word on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="688a4-162">Quando você executar este comando, o servidor da Web local será iniciado (se ainda não estiver em execução).</span><span class="sxs-lookup"><span data-stu-id="688a4-162">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="688a4-163">Para usar o seu suplemento, abra um novo documento no Word na Web e em seguida realize o sideload no suplemento de acordo com as instruções em [Realizar Sideload nos Suplementos do Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="688a4-163">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

2. <span data-ttu-id="688a4-164">No Word, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na Faixa de Opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="688a4-164">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Uma captura de tela do aplicativo Word com o botão Mostrar Painel de Tarefas realçado](../images/word-quickstart-addin-2b.png)

3. <span data-ttu-id="688a4-166">No painel de tarefas, escolha o botão **Inserir Parágrafo**.</span><span class="sxs-lookup"><span data-stu-id="688a4-166">In the task pane, choose the **Insert Paragraph** button.</span></span>

4. <span data-ttu-id="688a4-167">Faça uma alteração no parágrafo.</span><span class="sxs-lookup"><span data-stu-id="688a4-167">Make a change in the paragraph.</span></span>

5. <span data-ttu-id="688a4-168">Escolha novamente o botão **Inserir Parágrafo**.</span><span class="sxs-lookup"><span data-stu-id="688a4-168">Choose the **Insert Paragraph** button again.</span></span> <span data-ttu-id="688a4-169">Observe que o novo parágrafo está acima do anterior porque o método `insertParagraph` está inserido no início do corpo do documento.</span><span class="sxs-lookup"><span data-stu-id="688a4-169">Note that the new paragraph appears above the previous one because the `insertParagraph` method is inserting at the start of the document's body.</span></span>

    ![Tutorial do Word: Inserir Parágrafo](../images/word-tutorial-insert-paragraph-2.png)

## <a name="format-text"></a><span data-ttu-id="688a4-171">Formatar texto</span><span class="sxs-lookup"><span data-stu-id="688a4-171">Format text</span></span>

<span data-ttu-id="688a4-172">Nesta etapa do tutorial, você aplicará um estilo interno ao texto, aplicará um estilo personalizado ao texto e alterará a fonte do texto.</span><span class="sxs-lookup"><span data-stu-id="688a4-172">In this step of the tutorial, you'll apply a built-in style to text, apply a custom style to text, and change the font of text.</span></span>

### <a name="apply-a-built-in-style-to-text"></a><span data-ttu-id="688a4-173">Aplicar um estilo interno ao texto</span><span class="sxs-lookup"><span data-stu-id="688a4-173">Apply a built-in style to text</span></span>

1. <span data-ttu-id="688a4-174">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="688a4-174">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="688a4-175">Localize o elemento `<button>` para o botão `insert-paragraph` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="688a4-175">Locate the `<button>` element for the `insert-paragraph` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="apply-style">Apply Style</button><br/><br/>
    ```

3. <span data-ttu-id="688a4-176">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-176">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="688a4-177">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `insert-paragraph` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="688a4-177">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `insert-paragraph` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("apply-style").onclick = applyStyle;
    ```

5. <span data-ttu-id="688a4-178">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="688a4-178">Add the following function to the end of the file:</span></span>

    ```js
    function applyStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to style text.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

6. <span data-ttu-id="688a4-179">Na função `applyStyle()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-179">Within the `applyStyle()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="688a4-180">O código aplica um estilo a um parágrafo, mas também é possível aplicar estilos em intervalos de texto.</span><span class="sxs-lookup"><span data-stu-id="688a4-180">Note that the code applies a style to a paragraph, but styles can also be applied to ranges of text.</span></span>

    ```js
    var firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

### <a name="apply-a-custom-style-to-text"></a><span data-ttu-id="688a4-181">Aplicar um estilo personalizado ao texto</span><span class="sxs-lookup"><span data-stu-id="688a4-181">Apply a custom style to text</span></span>

1. <span data-ttu-id="688a4-182">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="688a4-182">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="688a4-183">Localize o elemento `<button>` para o botão `apply-style` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="688a4-183">Locate the `<button>` element for the `apply-style` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button><br/><br/>
    ```

3. <span data-ttu-id="688a4-184">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-184">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="688a4-185">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `apply-style` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="688a4-185">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `apply-style` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("apply-custom-style").onclick = applyCustomStyle;
    ```

5. <span data-ttu-id="688a4-186">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="688a4-186">Add the following function to the end of the file:</span></span>

    ```js
    function applyCustomStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply the custom style.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

6. <span data-ttu-id="688a4-187">Na função `applyCustomStyle()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-187">Within the `applyCustomStyle()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="688a4-188">O código aplica um estilo personalizado que ainda não existe.</span><span class="sxs-lookup"><span data-stu-id="688a4-188">Note that the code applies a custom style that does not exist yet.</span></span> <span data-ttu-id="688a4-189">Você criará um estilo com o nome **MyCustomStyle** na etapa [Testar o suplemento](#test-the-add-in-1).</span><span class="sxs-lookup"><span data-stu-id="688a4-189">You'll create a style with the name **MyCustomStyle** in the [Test the add-in](#test-the-add-in-1) step.</span></span>

    ```js
    var lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

7. <span data-ttu-id="688a4-190">Verifique se você salvou todas as alterações feitas no projeto.</span><span class="sxs-lookup"><span data-stu-id="688a4-190">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="change-the-font-of-text"></a><span data-ttu-id="688a4-191">Alterar a fonte do texto</span><span class="sxs-lookup"><span data-stu-id="688a4-191">Change the font of text</span></span>

1. <span data-ttu-id="688a4-192">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="688a4-192">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="688a4-193">Localize o elemento `<button>` para o botão `apply-custom-style` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="688a4-193">Locate the `<button>` element for the `apply-custom-style` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="change-font">Change Font</button><br/><br/>
    ```

3. <span data-ttu-id="688a4-194">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-194">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="688a4-195">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `apply-custom-style` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="688a4-195">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `apply-custom-style` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("change-font").onclick = changeFont;
    ```

5. <span data-ttu-id="688a4-196">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="688a4-196">Add the following function to the end of the file:</span></span>

    ```js
    function changeFont() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply a different font.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

6. <span data-ttu-id="688a4-197">Na função `changeFont()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-197">Within the `changeFont()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="688a4-198">O código recebe uma referência para o segundo parágrafo usando o método `ParagraphCollection.getFirst` encadeado para o método `Paragraph.getNext`.</span><span class="sxs-lookup"><span data-stu-id="688a4-198">Note that the code gets a reference to the second paragraph by using the `ParagraphCollection.getFirst` method chained to the `Paragraph.getNext` method.</span></span>

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

7. <span data-ttu-id="688a4-199">Verifique se você salvou todas as alterações feitas no projeto.</span><span class="sxs-lookup"><span data-stu-id="688a4-199">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="688a4-200">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="688a4-200">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

2. <span data-ttu-id="688a4-201">Se o painel de tarefas do suplemento ainda não estiver aberto no Word, vá para a guia **Página Inicial** e escolha o botão **Mostrar Painel de Tarefas** na faixa de opções para abri-lo.</span><span class="sxs-lookup"><span data-stu-id="688a4-201">If the add-in task pane isn't already open in Word, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="688a4-202">Verifique se há pelo menos três parágrafos no documento.</span><span class="sxs-lookup"><span data-stu-id="688a4-202">Be sure there are at least three paragraphs in the document.</span></span> <span data-ttu-id="688a4-203">É possível escolher o botão **Inserir Parágrafo** três vezes.</span><span class="sxs-lookup"><span data-stu-id="688a4-203">You can choose the **Insert Paragraph** button three times.</span></span> <span data-ttu-id="688a4-204">*Verifique com atenção se não há um parágrafo em branco no final do documento. Se houver, exclua-o.*</span><span class="sxs-lookup"><span data-stu-id="688a4-204">*Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.*</span></span>

4. <span data-ttu-id="688a4-205">No Word, crie um [estilo personalizado](https://support.office.com/article/customize-or-create-new-styles-d38d6e47-f6fc-48eb-a607-1eb120dec563) chamado de "MyCustomStyle".</span><span class="sxs-lookup"><span data-stu-id="688a4-205">In Word, create a [custom style](https://support.office.com/article/customize-or-create-new-styles-d38d6e47-f6fc-48eb-a607-1eb120dec563) named "MyCustomStyle".</span></span> <span data-ttu-id="688a4-206">Pode ter a formatação que você quiser.</span><span class="sxs-lookup"><span data-stu-id="688a4-206">It can have any formatting that you want.</span></span>

5. <span data-ttu-id="688a4-207">Choose the **Apply Style** button.</span><span class="sxs-lookup"><span data-stu-id="688a4-207">Choose the **Apply Style** button.</span></span> <span data-ttu-id="688a4-208">The first paragraph will be styled with the built-in style **Intense Reference**.</span><span class="sxs-lookup"><span data-stu-id="688a4-208">The first paragraph will be styled with the built-in style **Intense Reference**.</span></span>

6. <span data-ttu-id="688a4-209">Choose the **Apply Custom Style** button.</span><span class="sxs-lookup"><span data-stu-id="688a4-209">Choose the **Apply Custom Style** button.</span></span> <span data-ttu-id="688a4-210">The last paragraph will be styled with your custom style.</span><span class="sxs-lookup"><span data-stu-id="688a4-210">The last paragraph will be styled with your custom style.</span></span> <span data-ttu-id="688a4-211">(If nothing seems to happen, the last paragraph might be blank.</span><span class="sxs-lookup"><span data-stu-id="688a4-211">(If nothing seems to happen, the last paragraph might be blank.</span></span> <span data-ttu-id="688a4-212">If so, add some text to it.)</span><span class="sxs-lookup"><span data-stu-id="688a4-212">If so, add some text to it.)</span></span>

7. <span data-ttu-id="688a4-213">Choose the **Change Font** button.</span><span class="sxs-lookup"><span data-stu-id="688a4-213">Choose the **Change Font** button.</span></span> <span data-ttu-id="688a4-214">The font of the second paragraph changes to 18 pt., bold, Courier New.</span><span class="sxs-lookup"><span data-stu-id="688a4-214">The font of the second paragraph changes to 18 pt., bold, Courier New.</span></span>

    ![Tutorial do Word: Aplicar estilos e fonte](../images/word-tutorial-apply-styles-and-font-2.png)

## <a name="replace-text-and-insert-text"></a><span data-ttu-id="688a4-216">Substituir texto e inserir texto</span><span class="sxs-lookup"><span data-stu-id="688a4-216">Replace text and insert text</span></span>

<span data-ttu-id="688a4-217">Nesta etapa o tutorial, você adicionará texto dentro e fora dos intervalos de texto selecionados, e substituirá o texto de um intervalo selecionado.</span><span class="sxs-lookup"><span data-stu-id="688a4-217">In this step of the tutorial, you'll add text inside and outside of selected ranges of text, and replace the text of a selected range.</span></span>

### <a name="add-text-inside-a-range"></a><span data-ttu-id="688a4-218">Adicionar texto dentro de um intervalo</span><span class="sxs-lookup"><span data-stu-id="688a4-218">Add text inside a range</span></span>

1. <span data-ttu-id="688a4-219">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="688a4-219">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="688a4-220">Localize o elemento `<button>` para o botão `change-font` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="688a4-220">Locate the `<button>` element for the `change-font` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button><br/><br/>
    ```

3. <span data-ttu-id="688a4-221">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-221">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="688a4-222">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `change-font` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="688a4-222">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `change-font` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
    ```
5. <span data-ttu-id="688a4-223">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="688a4-223">Add the following function to the end of the file:</span></span>

    ```js
    function insertTextIntoRange() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert text into a selected range.

            // TODO2: Load the text of the range and sync so that the
            //        current range text can be read.

            // TODO3: Queue commands to repeat the text of the original
            //        range at the end of the document.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

6. <span data-ttu-id="688a4-224">Na função `insertTextIntoRange()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-224">Within the `insertTextIntoRange()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="688a4-225">Observação:</span><span class="sxs-lookup"><span data-stu-id="688a4-225">Note:</span></span>

   - <span data-ttu-id="688a4-226">The method is intended to insert the abbreviation ["(C2R)"] into the end of the Range whose text is "Click-to-Run".</span><span class="sxs-lookup"><span data-stu-id="688a4-226">The method is intended to insert the abbreviation ["(C2R)"] into the end of the Range whose text is "Click-to-Run".</span></span> <span data-ttu-id="688a4-227">It makes a simplifying assumption that the string is present and the user has selected it.</span><span class="sxs-lookup"><span data-stu-id="688a4-227">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="688a4-228">O primeiro parâmetro do método `Range.insertText` é a cadeia de caracteres a ser inserida no objeto `Range`.</span><span class="sxs-lookup"><span data-stu-id="688a4-228">The first parameter of the `Range.insertText` method is the string to insert into the `Range` object.</span></span>

   - <span data-ttu-id="688a4-229">The second parameter specifies where in the range the additional text should be inserted.</span><span class="sxs-lookup"><span data-stu-id="688a4-229">The second parameter specifies where in the range the additional text should be inserted.</span></span> <span data-ttu-id="688a4-230">Besides "End", the other possible options are "Start", "Before", "After", and "Replace".</span><span class="sxs-lookup"><span data-stu-id="688a4-230">Besides "End", the other possible options are "Start", "Before", "After", and "Replace".</span></span> 

   - <span data-ttu-id="688a4-231">The difference between "End" and "After" is that "End" inserts the new text inside the end of the existing range, but "After" creates a new range with the string and inserts the new range after the existing range.</span><span class="sxs-lookup"><span data-stu-id="688a4-231">The difference between "End" and "After" is that "End" inserts the new text inside the end of the existing range, but "After" creates a new range with the string and inserts the new range after the existing range.</span></span> <span data-ttu-id="688a4-232">Similarly, "Start" inserts text inside the beginning of the existing range and "Before" inserts a new range.</span><span class="sxs-lookup"><span data-stu-id="688a4-232">Similarly, "Start" inserts text inside the beginning of the existing range and "Before" inserts a new range.</span></span> <span data-ttu-id="688a4-233">"Replace" replaces the text of the existing range with the string in the first parameter.</span><span class="sxs-lookup"><span data-stu-id="688a4-233">"Replace" replaces the text of the existing range with the string in the first parameter.</span></span>

   - <span data-ttu-id="688a4-234">You saw in an earlier stage of the tutorial that the insert\* methods of the body object do not have the "Before" and "After" options.</span><span class="sxs-lookup"><span data-stu-id="688a4-234">You saw in an earlier stage of the tutorial that the insert\* methods of the body object do not have the "Before" and "After" options.</span></span> <span data-ttu-id="688a4-235">This is because you can't put content outside of the document's body.</span><span class="sxs-lookup"><span data-stu-id="688a4-235">This is because you can't put content outside of the document's body.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ```

7. <span data-ttu-id="688a4-236">Vamos deixar `TODO2` de lado até a próxima seção.</span><span class="sxs-lookup"><span data-stu-id="688a4-236">We'll skip over `TODO2` until the next section.</span></span> <span data-ttu-id="688a4-237">Na função `insertTextIntoRange()`, substitua `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-237">Within the `insertTextIntoRange()` function, replace `TODO3` with the following code.</span></span> <span data-ttu-id="688a4-238">Esse código é semelhante ao código que você criou no primeiro estágio do tutorial, exceto que, agora, você está inserindo um novo parágrafo no final do documento, em vez de no início.</span><span class="sxs-lookup"><span data-stu-id="688a4-238">This code is similar to the code you created in the first stage of the tutorial, except that now you are inserting a new paragraph at the end of the document instead of at the start.</span></span> <span data-ttu-id="688a4-239">Este novo parágrafo demonstrará que o novo texto agora faz parte do intervalo original.</span><span class="sxs-lookup"><span data-stu-id="688a4-239">This new paragraph will demonstrate that the new text is now part of the original range.</span></span>

    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text, "End");
    ```

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="688a4-240">Adicione código para buscar propriedades do documento em objetos de script do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="688a4-240">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="688a4-241">Em todas as funções anteriores desta série de tutoriais, você colocou em fila comandos para *gravar* no documento do Office.</span><span class="sxs-lookup"><span data-stu-id="688a4-241">In all previous functions in this series of tutorials, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="688a4-242">Cada função terminou com uma chamada para o método `context.sync()`, que envia os comandos em fila para o documento a ser executado.</span><span class="sxs-lookup"><span data-stu-id="688a4-242">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="688a4-243">Entretanto, o código adicionado na última etapa chama a propriedade `originalRange.text` e essa é uma grande diferença das funções anteriores que você escreveu, pois o objeto `originalRange` é apenas um objeto de proxy que existe no script do seu painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="688a4-243">But the code you added in the last step calls the `originalRange.text` property, and this is a significant difference from the earlier functions you wrote, because the `originalRange` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="688a4-244">Ele não sabe qual é o texto real do intervalo no documento, portanto, sua propriedade `text` não pode ter um valor real.</span><span class="sxs-lookup"><span data-stu-id="688a4-244">It doesn't know what the actual text of the range in the document is, so its `text` property can't have a real value.</span></span> <span data-ttu-id="688a4-245">Primeiro, é necessário buscar o valor de texto do intervalo no documento e usá-lo para definir o valor de `originalRange.text`.</span><span class="sxs-lookup"><span data-stu-id="688a4-245">It is necessary to first fetch the text value of the range from the document and use it to set the value of `originalRange.text`.</span></span> <span data-ttu-id="688a4-246">Somente então será possível chamar `originalRange.text` sem causar uma exceção.</span><span class="sxs-lookup"><span data-stu-id="688a4-246">Only then can `originalRange.text` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="688a4-247">Esse processo de busca tem três etapas:</span><span class="sxs-lookup"><span data-stu-id="688a4-247">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="688a4-248">Coloque em fila um comando para carregar (ou seja, fetch) as propriedades que seu código precisa ler.</span><span class="sxs-lookup"><span data-stu-id="688a4-248">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="688a4-249">Chame o método `sync` do objeto de contexto para enviar o comando em fila para o documento para execução e retornar as informações solicitadas.</span><span class="sxs-lookup"><span data-stu-id="688a4-249">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="688a4-250">Como o método `sync` é assíncrono, certifique-se de que ele tenha sido concluído antes que o código chame as propriedades que foram buscadas.</span><span class="sxs-lookup"><span data-stu-id="688a4-250">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="688a4-251">Essas etapas devem ser concluídas sempre que seu código precisar *ler* informações do documento do Office.</span><span class="sxs-lookup"><span data-stu-id="688a4-251">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="688a4-252">Na função `insertTextIntoRange()`, substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-252">Within the `insertTextIntoRange()` function, replace `TODO2` with the following code.</span></span>
  
    ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {
            // TODO4: Move the doc.body.insertParagraph line here.
        })
        // TODO5: Move the final call of context.sync here and ensure
        //        that it does not run until the insertParagraph has
        //        been queued.
    ```

2. <span data-ttu-id="688a4-253">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Word.run`.</span><span class="sxs-lookup"><span data-stu-id="688a4-253">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Word.run`.</span></span> <span data-ttu-id="688a4-254">You'll add a new final `context.sync` later in this tutorial.</span><span class="sxs-lookup"><span data-stu-id="688a4-254">You'll add a new final `context.sync` later in this tutorial.</span></span>

3. <span data-ttu-id="688a4-255">Recorte a linha `doc.body.insertParagraph` e cole no lugar de `TODO4`.</span><span class="sxs-lookup"><span data-stu-id="688a4-255">Cut the `doc.body.insertParagraph` line and paste in place of `TODO4`.</span></span>

4. <span data-ttu-id="688a4-256">Replace `TODO5` with the following code.</span><span class="sxs-lookup"><span data-stu-id="688a4-256">Replace `TODO5` with the following code.</span></span> <span data-ttu-id="688a4-257">Note:</span><span class="sxs-lookup"><span data-stu-id="688a4-257">Note:</span></span>

   - <span data-ttu-id="688a4-258">Passar o método `sync` para uma função `then` garante que ele não seja executado até que a lógica `insertParagraph` tenha sido enfileirada.</span><span class="sxs-lookup"><span data-stu-id="688a4-258">Passing the `sync` method to a `then` function ensures that it does not run until the `insertParagraph` logic has been queued.</span></span>

   - <span data-ttu-id="688a4-259">O método `then` invoca qualquer função que é passada para ele e não é recomendável que `sync` seja chamado duas vezes, portanto, omita os "()" do fim de context.sync.</span><span class="sxs-lookup"><span data-stu-id="688a4-259">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so omit the "()" from the end of context.sync.</span></span>

    ```js
    .then(context.sync);
    ```

<span data-ttu-id="688a4-260">Quando terminar, a função inteira deve se parecer com o seguinte:</span><span class="sxs-lookup"><span data-stu-id="688a4-260">When you're done, the entire function should look like the following:</span></span>

```js
function insertTextIntoRange() {
    Word.run(function (context) {

        var doc = context.document;
        var originalRange = doc.getSelection();
        originalRange.insertText(" (C2R)", "End");

        originalRange.load("text");
        return context.sync()
            .then(function() {
                doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");
            })
            .then(context.sync);
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
```

### <a name="add-text-between-ranges"></a><span data-ttu-id="688a4-261">Adicionar texto entre intervalos</span><span class="sxs-lookup"><span data-stu-id="688a4-261">Add text between ranges</span></span>

1. <span data-ttu-id="688a4-262">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="688a4-262">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="688a4-263">Localize o elemento `<button>` para o botão `insert-text-into-range` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="688a4-263">Locate the `<button>` element for the `insert-text-into-range` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button><br/><br/>
    ```

3. <span data-ttu-id="688a4-264">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-264">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="688a4-265">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `insert-text-into-range` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="688a4-265">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `insert-text-into-range` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;
    ```

5. <span data-ttu-id="688a4-266">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="688a4-266">Add the following function to the end of the file:</span></span>

    ```js
    function insertTextBeforeRange() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert a new range before the
            //        selected range.

            // TODO2: Load the text of the original range and sync so that the
            //        range text can be read and inserted.

        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="688a4-267">Na função `insertTextBeforeRange()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-267">Within the `insertTextBeforeRange()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="688a4-268">Observação:</span><span class="sxs-lookup"><span data-stu-id="688a4-268">Note:</span></span>

   - <span data-ttu-id="688a4-269">The method is intended to add a range whose text is "Office 2019, " before the range with text "Office 365".</span><span class="sxs-lookup"><span data-stu-id="688a4-269">The method is intended to add a range whose text is "Office 2019, " before the range with text "Office 365".</span></span> <span data-ttu-id="688a4-270">It makes a simplifying assumption that the string is present and the user has selected it.</span><span class="sxs-lookup"><span data-stu-id="688a4-270">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="688a4-271">O primeiro parâmetro do método `Range.insertText` é a cadeia de caracteres a ser adicionada.</span><span class="sxs-lookup"><span data-stu-id="688a4-271">The first parameter of the `Range.insertText` method is the string to add.</span></span>

   - <span data-ttu-id="688a4-272">The second parameter specifies where in the range the additional text should be inserted.</span><span class="sxs-lookup"><span data-stu-id="688a4-272">The second parameter specifies where in the range the additional text should be inserted.</span></span> <span data-ttu-id="688a4-273">For more details about the location options, see the previous discussion of the `insertTextIntoRange` function.</span><span class="sxs-lookup"><span data-stu-id="688a4-273">For more details about the location options, see the previous discussion of the `insertTextIntoRange` function.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ```

7. <span data-ttu-id="688a4-274">Na função `insertTextBeforeRange()`, substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-274">Within the `insertTextBeforeRange()` function, replace `TODO2` with the following code.</span></span>

     ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {
            // TODO3: Queue commands to insert the original range as a
            //        paragraph at the end of the document.
        })
        // TODO4: Make a final call of context.sync here and ensure
        //        that it does not run until the insertParagraph has
        //        been queued.
    ```

8. <span data-ttu-id="688a4-275">Replace `TODO3` with the following code.</span><span class="sxs-lookup"><span data-stu-id="688a4-275">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="688a4-276">This new paragraph will demonstrate the fact that the new text is ***not*** part of the original selected range.</span><span class="sxs-lookup"><span data-stu-id="688a4-276">This new paragraph will demonstrate the fact that the new text is ***not*** part of the original selected range.</span></span> <span data-ttu-id="688a4-277">The original range still has only the text it had when it was selected.</span><span class="sxs-lookup"><span data-stu-id="688a4-277">The original range still has only the text it had when it was selected.</span></span>

    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");
    ```

9. <span data-ttu-id="688a4-278">Substitua `TODO4` pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="688a4-278">Replace `TODO4` with the following code:</span></span>

    ```js
    .then(context.sync);
    ```

### <a name="replace-the-text-of-a-range"></a><span data-ttu-id="688a4-279">Substitua o texto de um intervalo.</span><span class="sxs-lookup"><span data-stu-id="688a4-279">Replace the text of a range</span></span>

1. <span data-ttu-id="688a4-280">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="688a4-280">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="688a4-281">Localize o elemento `<button>` para o botão `insert-text-outside-range` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="688a4-281">Locate the `<button>` element for the `insert-text-outside-range` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="replace-text">Change Quantity Term</button><br/><br/>
    ```

3. <span data-ttu-id="688a4-282">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-282">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="688a4-283">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `insert-text-outside-range` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="688a4-283">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `insert-text-outside-range` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("replace-text").onclick = replaceText;
    ```

5. <span data-ttu-id="688a4-284">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="688a4-284">Add the following function to the end of the file:</span></span>

    ```js
    function replaceText() {
        Word.run(function (context) {

            // TODO1: Queue commands to replace the text.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="688a4-285">Na função `replaceText()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-285">Within the `replaceText()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="688a4-286">O método serve para substituir a cadeia de caracteres "várias" pela cadeia "muitos".</span><span class="sxs-lookup"><span data-stu-id="688a4-286">Note that the method is intended to replace the string "several" with the string "many".</span></span> <span data-ttu-id="688a4-287">Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.</span><span class="sxs-lookup"><span data-stu-id="688a4-287">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    ```

7. <span data-ttu-id="688a4-288">Verifique se você salvou todas as alterações feitas no projeto.</span><span class="sxs-lookup"><span data-stu-id="688a4-288">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="688a4-289">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="688a4-289">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

2. <span data-ttu-id="688a4-290">Se o painel de tarefas do suplemento ainda não estiver aberto no Word, vá para a guia **Página Inicial** e escolha o botão **Mostrar Painel de Tarefas** na faixa de opções para abri-lo.</span><span class="sxs-lookup"><span data-stu-id="688a4-290">If the add-in task pane isn't already open in Word, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="688a4-291">No painel de tarefas, escolha o botão **Inserir Parágrafo** para garantir que haja um parágrafo no início do documento.</span><span class="sxs-lookup"><span data-stu-id="688a4-291">In the task pane, choose the **Insert Paragraph** button to ensure that there is a paragraph at the start of the document.</span></span>

4. <span data-ttu-id="688a4-292">No documento, selecione a frase "Clique para Executar".</span><span class="sxs-lookup"><span data-stu-id="688a4-292">Within the document, select the phrase "Click-to-Run".</span></span> <span data-ttu-id="688a4-293">*Tenha cuidado para não incluir o espaço anterior ou a vírgula seguinte na seleção.*</span><span class="sxs-lookup"><span data-stu-id="688a4-293">*Be careful not to include the preceding space or following comma in the selection.*</span></span>

5. <span data-ttu-id="688a4-294">Choose the **Insert Abbreviation** button.</span><span class="sxs-lookup"><span data-stu-id="688a4-294">Choose the **Insert Abbreviation** button.</span></span> <span data-ttu-id="688a4-295">Note that " (C2R)" is added.</span><span class="sxs-lookup"><span data-stu-id="688a4-295">Note that " (C2R)" is added.</span></span> <span data-ttu-id="688a4-296">Note also that at the bottom of the document a new paragraph is added with the entire expanded text because the new string was added to the existing range.</span><span class="sxs-lookup"><span data-stu-id="688a4-296">Note also that at the bottom of the document a new paragraph is added with the entire expanded text because the new string was added to the existing range.</span></span>

6. <span data-ttu-id="688a4-297">No documento, selecione a frase "Office 365".</span><span class="sxs-lookup"><span data-stu-id="688a4-297">Within the document, select the phrase "Office 365".</span></span> <span data-ttu-id="688a4-298">*Tenha cuidado para não incluir o espaço anterior ou seguinte na seleção.*</span><span class="sxs-lookup"><span data-stu-id="688a4-298">*Be careful not to include the preceding or following space in the selection.*</span></span>

7. <span data-ttu-id="688a4-299">Choose the **Add Version Info** button.</span><span class="sxs-lookup"><span data-stu-id="688a4-299">Choose the **Add Version Info** button.</span></span> <span data-ttu-id="688a4-300">Note that "Office 2019, " is inserted between "Office 2016" and "Office 365".</span><span class="sxs-lookup"><span data-stu-id="688a4-300">Note that "Office 2019, " is inserted between "Office 2016" and "Office 365".</span></span> <span data-ttu-id="688a4-301">Note also that at the bottom of the document a new paragraph is added but it contains only the originally selected text because the new string became a new range rather than being added to the original range.</span><span class="sxs-lookup"><span data-stu-id="688a4-301">Note also that at the bottom of the document a new paragraph is added but it contains only the originally selected text because the new string became a new range rather than being added to the original range.</span></span>

8. <span data-ttu-id="688a4-302">No documento, selecione a palavra "vários".</span><span class="sxs-lookup"><span data-stu-id="688a4-302">Within the document, select the word "several".</span></span> <span data-ttu-id="688a4-303">*Tenha cuidado para não incluir o espaço anterior ou seguinte na seleção.*</span><span class="sxs-lookup"><span data-stu-id="688a4-303">*Be careful not to include the preceding or following space in the selection.*</span></span>

9. <span data-ttu-id="688a4-304">Choose the **Change Quantity Term** button.</span><span class="sxs-lookup"><span data-stu-id="688a4-304">Choose the **Change Quantity Term** button.</span></span> <span data-ttu-id="688a4-305">Note that "many" replaces the selected text.</span><span class="sxs-lookup"><span data-stu-id="688a4-305">Note that "many" replaces the selected text.</span></span>

    ![Tutorial do Word: texto adicionado e substituído](../images/word-tutorial-text-replace-2.png)

## <a name="insert-images-html-and-tables"></a><span data-ttu-id="688a4-307">Inserir imagens, HTML e tabelas</span><span class="sxs-lookup"><span data-stu-id="688a4-307">Insert images, HTML, and tables</span></span>

<span data-ttu-id="688a4-308">Nesta etapa do tutorial, você aprenderá a inserir imagens, HTML e tabelas no documento.</span><span class="sxs-lookup"><span data-stu-id="688a4-308">In this step of the tutorial, you'll learn how to insert images, HTML, and tables into the document.</span></span>

### <a name="define-an-image"></a><span data-ttu-id="688a4-309">Definir uma imagem</span><span class="sxs-lookup"><span data-stu-id="688a4-309">Define an image</span></span>

<span data-ttu-id="688a4-310">Conclua as seguintes etapas para definir a imagem que será inserida no documento na próxima parte deste tutorial.</span><span class="sxs-lookup"><span data-stu-id="688a4-310">Complete the following steps to define the image that you'll insert into the document in the next part of this tutorial.</span></span> 

1. <span data-ttu-id="688a4-311">Na raiz do projeto, crie um novo arquivo chamado **base64Image.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-311">In the root of the project, create a new file named **base64Image.js**.</span></span>

2. <span data-ttu-id="688a4-312">Abra o arquivo **base64Image.js** e adicione o seguinte código para especificar a cadeia de caracteres codificada em base64 que representa uma imagem.</span><span class="sxs-lookup"><span data-stu-id="688a4-312">Open the file **base64Image.js** and add the following code to specify the base64-encoded string that represents an image.</span></span>

    ```js
    export const base64Image =
        "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";
    ```

### <a name="insert-an-image"></a><span data-ttu-id="688a4-313">Inserir uma imagem</span><span class="sxs-lookup"><span data-stu-id="688a4-313">Insert an image</span></span>

1. <span data-ttu-id="688a4-314">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="688a4-314">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="688a4-315">Localize o elemento `<button>` para o botão `replace-text` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="688a4-315">Locate the `<button>` element for the `replace-text` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="insert-image">Insert Image</button><br/><br/>
    ```

3. <span data-ttu-id="688a4-316">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-316">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="688a4-317">Localize a chamada do método `Office.onReady`, próximo à parte superior e adicione o seguinte código imediatamente antes dessa linha.</span><span class="sxs-lookup"><span data-stu-id="688a4-317">Locate the `Office.onReady` method call near the top of the file and add the following code immediately before that line.</span></span> <span data-ttu-id="688a4-318">Esse código importa a variável que você definida anteriormente no arquivo **./base64Image.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-318">This code imports the variable that you defined previously in the file **./base64Image.js**.</span></span>

    ```js
    import { base64Image } from "../../base64Image";
    ```

5. <span data-ttu-id="688a4-319">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `replace-text` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="688a4-319">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `replace-text` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("insert-image").onclick = insertImage;
    ```

6. <span data-ttu-id="688a4-320">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="688a4-320">Add the following function to the end of the file:</span></span>

    ```js
    function insertImage() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert an image.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

7. <span data-ttu-id="688a4-321">Na função `insertImage()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-321">Within the `insertImage()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="688a4-322">Esta linha insere a imagem codificada em base 64 no final do documento.</span><span class="sxs-lookup"><span data-stu-id="688a4-322">Note that this line inserts the base 64 encoded image at the end of the document.</span></span> <span data-ttu-id="688a4-323">(O objeto `Paragraph` também tem um método `insertInlinePictureFromBase64` e outros métodos `insert*`.</span><span class="sxs-lookup"><span data-stu-id="688a4-323">(The `Paragraph` object also has an `insertInlinePictureFromBase64` method and other `insert*` methods.</span></span> <span data-ttu-id="688a4-324">Confira a seção insertHTML a seguir para conferir um exemplo).</span><span class="sxs-lookup"><span data-stu-id="688a4-324">See the following insertHTML section for an example.)</span></span>

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

### <a name="insert-html"></a><span data-ttu-id="688a4-325">Inserir HTML</span><span class="sxs-lookup"><span data-stu-id="688a4-325">Insert HTML</span></span>

1. <span data-ttu-id="688a4-326">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="688a4-326">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="688a4-327">Localize o elemento `<button>` para o botão `insert-image` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="688a4-327">Locate the `<button>` element for the `insert-image` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="insert-html">Insert HTML</button><br/><br/>
    ```

3. <span data-ttu-id="688a4-328">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-328">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="688a4-329">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `insert-image` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="688a4-329">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `insert-image` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("insert-html").onclick = insertHTML;
    ```
5. <span data-ttu-id="688a4-330">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="688a4-330">Add the following function to the end of the file:</span></span>

    ```js
    function insertHTML() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert a string of HTML.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="688a4-331">Na função `insertHTML()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-331">Within the `insertHTML()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="688a4-332">Observação:</span><span class="sxs-lookup"><span data-stu-id="688a4-332">Note:</span></span>

   - <span data-ttu-id="688a4-333">A primeira linha adiciona um parágrafo em branco ao final do documento.</span><span class="sxs-lookup"><span data-stu-id="688a4-333">The first line adds a blank paragraph to the end of the document.</span></span> 

   - <span data-ttu-id="688a4-334">The second line inserts a string of HTML at the end of the paragraph; specifically two paragraphs, one formatted with Verdana font, the other with the default styling of the Word document.</span><span class="sxs-lookup"><span data-stu-id="688a4-334">The second line inserts a string of HTML at the end of the paragraph; specifically two paragraphs, one formatted with Verdana font, the other with the default styling of the Word document.</span></span> <span data-ttu-id="688a4-335">(As you saw in the `insertImage` method earlier, the `context.document.body` object also has the `insert*` methods.)</span><span class="sxs-lookup"><span data-stu-id="688a4-335">(As you saw in the `insertImage` method earlier, the `context.document.body` object also has the `insert*` methods.)</span></span>

    ```js
    var blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

### <a name="insert-a-table"></a><span data-ttu-id="688a4-336">Inserir uma tabela</span><span class="sxs-lookup"><span data-stu-id="688a4-336">Insert a table</span></span>

1. <span data-ttu-id="688a4-337">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="688a4-337">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="688a4-338">Localize o elemento `<button>` para o botão `insert-html` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="688a4-338">Locate the `<button>` element for the `insert-html` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="insert-table">Insert Table</button><br/><br/>
    ```

3. <span data-ttu-id="688a4-339">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-339">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="688a4-340">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `insert-html` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="688a4-340">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `insert-html` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("insert-table").onclick = insertTable;
    ```

5. <span data-ttu-id="688a4-341">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="688a4-341">Add the following function to the end of the file:</span></span>

    ```js
    function insertTable() {
        Word.run(function (context) {

            // TODO1: Queue commands to get a reference to the paragraph
            //        that will proceed the table.

            // TODO2: Queue commands to create a table and populate it with data.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="688a4-342">Na função `insertTable()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-342">Within the `insertTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="688a4-343">Essa linha usa o método `ParagraphCollection.getFirst` para obter uma referência do primeiro parágrafo e, depois, usa o método `Paragraph.getNext` para obter uma referência para o segundo parágrafo.</span><span class="sxs-lookup"><span data-stu-id="688a4-343">Note that this line uses the `ParagraphCollection.getFirst` method to get a reference ot the first paragraph and then uses the `Paragraph.getNext` method to get a reference to the second paragraph.</span></span>

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

7. <span data-ttu-id="688a4-344">Na função `insertTable()`, substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-344">Within the `insertTable()` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="688a4-345">Observação:</span><span class="sxs-lookup"><span data-stu-id="688a4-345">Note:</span></span>

   - <span data-ttu-id="688a4-346">Os dois primeiros parâmetros do método `insertTable` especificam o número de linhas e colunas.</span><span class="sxs-lookup"><span data-stu-id="688a4-346">The first two parameters of the `insertTable` method specify the number of rows and columns.</span></span>

   - <span data-ttu-id="688a4-347">O terceiro parâmetro especifica onde inserir a tabela, nesse caso, depois do parágrafo.</span><span class="sxs-lookup"><span data-stu-id="688a4-347">The third parameter specifies where to insert the table, in this case after the paragraph.</span></span>

   - <span data-ttu-id="688a4-348">O quarto parâmetro é uma matriz bidimensional que define os valores das células da tabela.</span><span class="sxs-lookup"><span data-stu-id="688a4-348">The fourth parameter is a two-dimensional array that sets the values of the table cells.</span></span>

   - <span data-ttu-id="688a4-349">A tabela terá um estilo padrão simples, mas o método `insertTable` retornará um objeto `Table` com muitos membros, e alguns deles são usados para alterar o estilo de tabela.</span><span class="sxs-lookup"><span data-stu-id="688a4-349">The table will have plain default styling, but the `insertTable` method returns a `Table` object with many members, some of which are used to style the table.</span></span>

    ```js
    var tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

8. <span data-ttu-id="688a4-350">Verifique se você salvou todas as alterações feitas no projeto.</span><span class="sxs-lookup"><span data-stu-id="688a4-350">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="688a4-351">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="688a4-351">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

2. <span data-ttu-id="688a4-352">Se o painel de tarefas do suplemento ainda não estiver aberto no Word, vá para a guia **Página Inicial** e escolha o botão **Mostrar Painel de Tarefas** na faixa de opções para abri-lo.</span><span class="sxs-lookup"><span data-stu-id="688a4-352">If the add-in task pane isn't already open in Word, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="688a4-353">No painel de tarefas, escolha o botão **Inserir Parágrafo** pelo menos três vezes para garantir que haja alguns parágrafos no documento.</span><span class="sxs-lookup"><span data-stu-id="688a4-353">In the task pane, choose the **Insert Paragraph** button at least three times to ensure that there are a few paragraphs in the document.</span></span>

4. <span data-ttu-id="688a4-354">Escolha o botão **Inserir Imagem**. Uma imagem é inserida no final do documento.</span><span class="sxs-lookup"><span data-stu-id="688a4-354">Choose the **Insert Image** button and note that an image is inserted at the end of the document.</span></span>

5. <span data-ttu-id="688a4-355">Escolha o botão **Inserir HTML**. Dois parágrafos são inseridos no final do documento, e o primeiro tem a fonte Verdana.</span><span class="sxs-lookup"><span data-stu-id="688a4-355">Choose the **Insert HTML** button and note that two paragraphs are inserted at the end of the document, and that the first one has Verdana font.</span></span>

6. <span data-ttu-id="688a4-356">Escolha o botão **Inserir Tabela**. Uma tabela é inserida após o segundo parágrafo.</span><span class="sxs-lookup"><span data-stu-id="688a4-356">Choose the **Insert Table** button and note that a table is inserted after the second paragraph.</span></span>

    ![Tutorial do Word: Inserir imagem, HTML e tabela](../images/word-tutorial-insert-image-html-table-2.png)

## <a name="create-and-update-content-controls"></a><span data-ttu-id="688a4-358">Criar e atualizar os controles de conteúdo</span><span class="sxs-lookup"><span data-stu-id="688a4-358">Create and update content controls</span></span>

<span data-ttu-id="688a4-359">Nesta etapa do tutorial, você aprenderá a criar controles de conteúdo de Rich Text no documento e, depois, como inserir e substituir conteúdo nos controles.</span><span class="sxs-lookup"><span data-stu-id="688a4-359">In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.</span></span>

> [!NOTE]
> <span data-ttu-id="688a4-360">Há vários tipos de controles de conteúdo que podem ser adicionados a um documento do Word por meio da interface do usuário. Porém, no momento, só há suporte para controles de conteúdo de Rich Text no Word.js.</span><span class="sxs-lookup"><span data-stu-id="688a4-360">There are several types of content controls that can be added to a Word document through the UI, but currently only Rich Text content controls are supported by Word.js.</span></span>
>
> <span data-ttu-id="688a4-361">Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties.</span><span class="sxs-lookup"><span data-stu-id="688a4-361">Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties.</span></span> <span data-ttu-id="688a4-362">For details, see [Create forms that users complete or print in Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span><span class="sxs-lookup"><span data-stu-id="688a4-362">For details, see [Create forms that users complete or print in Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span></span>

### <a name="create-a-content-control"></a><span data-ttu-id="688a4-363">Criar um controle de conteúdo</span><span class="sxs-lookup"><span data-stu-id="688a4-363">Create a content control</span></span>

1. <span data-ttu-id="688a4-364">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="688a4-364">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="688a4-365">Localize o elemento `<button>` para o botão `insert-table` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="688a4-365">Locate the `<button>` element for the `insert-table` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="create-content-control">Create Content Control</button><br/><br/>
    ```

3. <span data-ttu-id="688a4-366">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-366">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="688a4-367">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `insert-table` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="688a4-367">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `insert-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("create-content-control").onclick = createContentControl;
    ```
5. <span data-ttu-id="688a4-368">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="688a4-368">Add the following function to the end of the file:</span></span>

    ```js
    function createContentControl() {
        Word.run(function (context) {

            // TODO1: Queue commands to create a content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="688a4-369">Na função `createContentControl()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-369">Within the `createContentControl()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="688a4-370">Observação:</span><span class="sxs-lookup"><span data-stu-id="688a4-370">Note:</span></span>

   - <span data-ttu-id="688a4-371">This code is intended to wrap the phrase "Office 365" in a content control.</span><span class="sxs-lookup"><span data-stu-id="688a4-371">This code is intended to wrap the phrase "Office 365" in a content control.</span></span> <span data-ttu-id="688a4-372">It makes a simplifying assumption that the string is present and the user has selected it.</span><span class="sxs-lookup"><span data-stu-id="688a4-372">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="688a4-373">A propriedade `ContentControl.title` especifica o título visível do controle de conteúdo.</span><span class="sxs-lookup"><span data-stu-id="688a4-373">The `ContentControl.title` property specifies the visible title of the content control.</span></span>

   - <span data-ttu-id="688a4-374">A propriedade `ContentControl.tag` especifica uma marca que pode ser usada para obter uma referência a um controle de conteúdo usando o método `ContentControlCollection.getByTag`, que você usará em uma função posterior.</span><span class="sxs-lookup"><span data-stu-id="688a4-374">The `ContentControl.tag` property specifies an tag that can be used to get a reference to a content control using the `ContentControlCollection.getByTag` method, which you'll use in a later function.</span></span>

   - <span data-ttu-id="688a4-375">The `ContentControl.appearance` property specifies the visual look of the control.</span><span class="sxs-lookup"><span data-stu-id="688a4-375">The `ContentControl.appearance` property specifies the visual look of the control.</span></span> <span data-ttu-id="688a4-376">Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title.</span><span class="sxs-lookup"><span data-stu-id="688a4-376">Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title.</span></span> <span data-ttu-id="688a4-377">Other possible values are "BoundingBox" and "None".</span><span class="sxs-lookup"><span data-stu-id="688a4-377">Other possible values are "BoundingBox" and "None".</span></span>

   - <span data-ttu-id="688a4-378">A propriedade `ContentControl.color` especifica a cor das marcas ou da borda da caixa delimitadora.</span><span class="sxs-lookup"><span data-stu-id="688a4-378">The `ContentControl.color` property specifies the color of the tags or the border of the bounding box.</span></span>

    ```js
    var serviceNameRange = context.document.getSelection();
    var serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ```

### <a name="replace-the-content-of-the-content-control"></a><span data-ttu-id="688a4-379">Substituir o conteúdo do controle de conteúdo</span><span class="sxs-lookup"><span data-stu-id="688a4-379">Replace the content of the content control</span></span>

1. <span data-ttu-id="688a4-380">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="688a4-380">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="688a4-381">Localize o elemento `<button>` para o botão `create-content-control` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="688a4-381">Locate the `<button>` element for the `create-content-control` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="replace-content-in-control">Rename Service</button><br/><br/>
    ```

3. <span data-ttu-id="688a4-382">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="688a4-382">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="688a4-383">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `create-content-control` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="688a4-383">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `create-content-control` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("replace-content-in-control").onclick = replaceContentInControl;
    ```

5. <span data-ttu-id="688a4-384">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="688a4-384">Add the following function to the end of the file:</span></span>

    ```js
    function replaceContentInControl() {
        Word.run(function (context) {

            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="688a4-385">Na função `replaceContentInControl()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="688a4-385">Within the `replaceContentInControl()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="688a4-386">Observação:</span><span class="sxs-lookup"><span data-stu-id="688a4-386">Note:</span></span>

    - <span data-ttu-id="688a4-387">O método `ContentControlCollection.getByTag` retorna um `ContentControlCollection` de todos os controles de conteúdo da marca especificada.</span><span class="sxs-lookup"><span data-stu-id="688a4-387">The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag.</span></span> <span data-ttu-id="688a4-388">Usamos `getFirst` para obter uma referência do controle desejado.</span><span class="sxs-lookup"><span data-stu-id="688a4-388">We use `getFirst` to get a reference to the desired control.</span></span>

    ```js
    var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ```

7. <span data-ttu-id="688a4-389">Verifique se você salvou todas as alterações feitas no projeto.</span><span class="sxs-lookup"><span data-stu-id="688a4-389">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="688a4-390">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="688a4-390">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

2. <span data-ttu-id="688a4-391">Se o painel de tarefas do suplemento ainda não estiver aberto no Word, vá para a guia **Página Inicial** e escolha o botão **Mostrar Painel de Tarefas** na faixa de opções para abri-lo.</span><span class="sxs-lookup"><span data-stu-id="688a4-391">If the add-in task pane isn't already open in Word, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="688a4-392">No painel de tarefas, escolha o botão **Inserir Parágrafo** para garantir que haja um parágrafo com "Office 365" no início do documento.</span><span class="sxs-lookup"><span data-stu-id="688a4-392">In the task pane, choose the **Insert Paragraph** button to ensure that there is a paragraph with "Office 365" at the top of the document.</span></span>

4. <span data-ttu-id="688a4-393">No documento, selecione o texto "Office 365" e, em seguida, escolha o botão **Criar Controle de Conteúdo**.</span><span class="sxs-lookup"><span data-stu-id="688a4-393">In the document, select the text "Office 365" and then choose the **Create Content Control** button.</span></span> <span data-ttu-id="688a4-394">A frase está envolvida por marcas chamadas "Nome do Serviço".</span><span class="sxs-lookup"><span data-stu-id="688a4-394">Note that the phrase is wrapped in tags labelled "Service Name".</span></span>

7. <span data-ttu-id="688a4-395">Escolha o botão **Renomear Serviço**. O texto do controle de conteúdo muda para "Fabrikam Online Productivity Suite".</span><span class="sxs-lookup"><span data-stu-id="688a4-395">Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".</span></span>

    ![Tutorial do Word - Criar o controle de conteúdo e alterar seu texto](../images/word-tutorial-content-control-2.png)

## <a name="next-steps"></a><span data-ttu-id="688a4-397">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="688a4-397">Next steps</span></span>

<span data-ttu-id="688a4-398">Neste tutorial, você criou um suplemento do painel de tarefas do Word que insere e substitui texto, imagens e outro conteúdo em um documento do Word.</span><span class="sxs-lookup"><span data-stu-id="688a4-398">In this tutorial, you've created a Word task pane add-in that inserts and replaces text, images, and other content in a Word document.</span></span> <span data-ttu-id="688a4-399">Para saber mais sobre o desenvolvimento de suplementos do Word, continue no seguinte artigo:</span><span class="sxs-lookup"><span data-stu-id="688a4-399">To learn more about building Word add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="688a4-400">Visão geral dos suplementos do Word</span><span class="sxs-lookup"><span data-stu-id="688a4-400">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="688a4-401">Confira também</span><span class="sxs-lookup"><span data-stu-id="688a4-401">See also</span></span>

* [<span data-ttu-id="688a4-402">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="688a4-402">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="688a4-403">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="688a4-403">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="688a4-404">Desenvolver Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="688a4-404">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
