---
title: Tutorial de suplemento do Word
description: Neste tutorial, voc? criar? um suplemento do Word que insere (e substitui) intervalos de texto, par?grafos, imagens, HTML, tabelas e controles de conte?do. Você também aprenderá como formatar texto e como inserir (e substituir) conteúdo nos controles de conteúdo.
ms.date: 12/31/2018
ms.topic: tutorial
ms.openlocfilehash: d1d278d1acd9e8a1377773b90ae9d528af69b93c
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724935"
---
# <a name="tutorial-create-a-word-task-pane-add-in"></a><span data-ttu-id="11c76-104">Tutorial: Criar Suplemento do Painel de Tarefas no Word</span><span class="sxs-lookup"><span data-stu-id="11c76-104">Create a dictionary task pane add-in</span></span>

<span data-ttu-id="11c76-105">Neste tutorial: você criará um suplemento do painel de tarefas no Word:</span><span class="sxs-lookup"><span data-stu-id="11c76-105">In this tutorial, you'll create a Word task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="11c76-106">Insere um intervalo de texto</span><span class="sxs-lookup"><span data-stu-id="11c76-106">Inserts a range of text</span></span>
> * <span data-ttu-id="11c76-107">Formatos de texto</span><span class="sxs-lookup"><span data-stu-id="11c76-107">Formats text</span></span>
> * <span data-ttu-id="11c76-108">Substitui e insere texto em vários locais</span><span class="sxs-lookup"><span data-stu-id="11c76-108">Replace text and insert text in various locations</span></span>
> * <span data-ttu-id="11c76-109">Insere imagens, HTML e tabelas</span><span class="sxs-lookup"><span data-stu-id="11c76-109">Insert images, HTML, and tables</span></span>
> * <span data-ttu-id="11c76-110">Cria e atualiza os controles de conteúdo</span><span class="sxs-lookup"><span data-stu-id="11c76-110">Creates and updates content controls</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="11c76-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="11c76-111">Prerequisites</span></span>

<span data-ttu-id="11c76-112">Para usar este tutorial, você precisa instalar o seguinte.</span><span class="sxs-lookup"><span data-stu-id="11c76-112">To use this tutorial, you need to have the following installed.</span></span> 

- <span data-ttu-id="11c76-113">Word 2016, versão 1711 (build 8730.1000 do Clique para Executar) ou posterior.</span><span class="sxs-lookup"><span data-stu-id="11c76-113">Word 2016, version 1711 (Build 8730.1000 Click-to-Run) or later.</span></span> <span data-ttu-id="11c76-114">Talvez você precise ser um participante do programa Office Insider para ter essa versão.</span><span class="sxs-lookup"><span data-stu-id="11c76-114">You might need to be an Office Insider to get this version.</span></span> <span data-ttu-id="11c76-115">Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="11c76-115">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

- [<span data-ttu-id="11c76-116">Nó</span><span class="sxs-lookup"><span data-stu-id="11c76-116">Node</span></span>](https://nodejs.org/en/) 

- <span data-ttu-id="11c76-117">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="11c76-117">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

## <a name="create-your-add-in-project"></a><span data-ttu-id="11c76-118">Criar seu projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="11c76-118">Create your add-in project</span></span>

<span data-ttu-id="11c76-119">Conclua as etapas a seguir para criar o projeto de suplemento do Word que você vai usar como base para este tutorial.</span><span class="sxs-lookup"><span data-stu-id="11c76-119">Complete the following steps to create the Word add-in project that you'll use as the basis for this tutorial.</span></span>

1. <span data-ttu-id="11c76-120">Clone o repositório do GitHub com o [Tutorial de suplemento do Word](https://github.com/OfficeDev/Word-Add-in-Tutorial).</span><span class="sxs-lookup"><span data-stu-id="11c76-120">Clone the GitHub repository [Word Add-in Tutorial](https://github.com/OfficeDev/Word-Add-in-Tutorial).</span></span>

2. <span data-ttu-id="11c76-121">Abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="11c76-121">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

3. <span data-ttu-id="11c76-122">Execute o comando `npm install` para instalar as ferramentas e bibliotecas listadas no arquivo package.json.</span><span class="sxs-lookup"><span data-stu-id="11c76-122">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 

4. <span data-ttu-id="11c76-123">Execute as etapas em [Adicionar certificados autoassinados como certificado raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para confiar no certificado do sistema operacional do seu computador de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="11c76-123">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

## <a name="insert-a-range-of-text"></a><span data-ttu-id="11c76-124">Inserir um intervalo de texto</span><span class="sxs-lookup"><span data-stu-id="11c76-124">Insert a range of cells</span></span>

<span data-ttu-id="11c76-125">Nesta etapa do tutorial, você testará programaticamente se o suplemento oferece suporte à versão atual do Word do usuário e inserirá um parágrafo no documento.</span><span class="sxs-lookup"><span data-stu-id="11c76-125">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Word, and then insert a paragraph in the document.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="11c76-126">Codificação do suplemento</span><span class="sxs-lookup"><span data-stu-id="11c76-126">Code the add-in</span></span>

1. <span data-ttu-id="11c76-127">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="11c76-127">Open the project in your code editor.</span></span>

2. <span data-ttu-id="11c76-128">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="11c76-128">Open the file index.html.</span></span>

3. <span data-ttu-id="11c76-129">Substitua `TODO1` pela marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-129">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. <span data-ttu-id="11c76-130">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="11c76-130">Open the app.js file.</span></span>

5. <span data-ttu-id="11c76-131">Substitua o `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="11c76-131">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="11c76-132">O código determina se a versão do Word do usuário suporta uma versão do Word.js que inclui todas as APIs usadas em todos os estágios deste tutorial.dae</span><span class="sxs-lookup"><span data-stu-id="11c76-132">This code determines whether the user's version of Word supports a version of Word.js that includes all the APIs that are used in all the stages of this tutorial.</span></span> <span data-ttu-id="11c76-133">Em um suplemento de produção, use o corpo do bloco condicional para ocultar ou desabilitar a interface do usuário que chame a APIs sem suporte.</span><span class="sxs-lookup"><span data-stu-id="11c76-133">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="11c76-134">Isso permitirá que o usuário ainda use as partes do suplemento às quais a versão do Word dá suporte.</span><span class="sxs-lookup"><span data-stu-id="11c76-134">This will enable the user to still use the parts of the add-in that are supported by their version of Word.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="11c76-135">Substitua o `TODO2` pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-135">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. <span data-ttu-id="11c76-136">Substitua o `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="11c76-136">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="11c76-137">Observação:</span><span class="sxs-lookup"><span data-stu-id="11c76-137">Note:</span></span>

   - <span data-ttu-id="11c76-138">A lógica de negócios de Word.js será adicionada à função que passar por `Word.run`.</span><span class="sxs-lookup"><span data-stu-id="11c76-138">Your Word.js business logic will be added to the function that is passed to `Word.run`.</span></span> <span data-ttu-id="11c76-139">Essa lógica não é executada imediatamente.</span><span class="sxs-lookup"><span data-stu-id="11c76-139">This logic does not execute immediately.</span></span> <span data-ttu-id="11c76-140">Em vez disso, ela é adicionada à fila de comandos pendentes.</span><span class="sxs-lookup"><span data-stu-id="11c76-140">Instead, it is added to a queue of pending commands.</span></span>

   - <span data-ttu-id="11c76-141">O método `context.sync` envia todos os comandos da fila para execução no Word.</span><span class="sxs-lookup"><span data-stu-id="11c76-141">The `context.sync` method sends all queued commands to Word for execution.</span></span>

   - <span data-ttu-id="11c76-142">O `Word.run` é seguido por um bloco `catch`.</span><span class="sxs-lookup"><span data-stu-id="11c76-142">The `Word.run` is followed by a `catch` block.</span></span> <span data-ttu-id="11c76-143">Essa é uma prática recomendada que você sempre deve seguir.</span><span class="sxs-lookup"><span data-stu-id="11c76-143">This is a best practice that you should always follow.</span></span> 

    ```js
    function insertParagraph() {
        Word.run(function (context) {

            // TODO4: Queue commands to insert a paragraph into the document.

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

8. <span data-ttu-id="11c76-p107">Substitua `TODO4` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="11c76-p107">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="11c76-146">O primeiro parâmetro para o método `insertParagraph` é o texto para o novo parágrafo.</span><span class="sxs-lookup"><span data-stu-id="11c76-146">The first parameter to the `insertParagraph` method is the text for the new paragraph.</span></span>

   - <span data-ttu-id="11c76-147">O segundo parâmetro é o local dentro do corpo onde o parágrafo será inserido.</span><span class="sxs-lookup"><span data-stu-id="11c76-147">The second parameter is the location within the body where the paragraph will be inserted.</span></span> <span data-ttu-id="11c76-148">Outras opções para inserir parágrafo, quando o objeto pai é o corpo, são "End" e "Replace".</span><span class="sxs-lookup"><span data-stu-id="11c76-148">Other options for insert paragraph, when the parent object is the body, are "End" and "Replace".</span></span>

    ```js
    var docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office Online.",
                            "Start");
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="11c76-149">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="11c76-149">Test the add-in</span></span>

1. <span data-ttu-id="11c76-150">Abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="11c76-150">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

2. <span data-ttu-id="11c76-151">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.</span><span class="sxs-lookup"><span data-stu-id="11c76-151">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="11c76-152">Execute o comando `npm start` para iniciar um servidor Web em um localhost.</span><span class="sxs-lookup"><span data-stu-id="11c76-152">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="11c76-153">Realize o sideload do suplemento usando um dos métodos a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-153">Sideload the add-in by using one of the following methods:</span></span>

    - <span data-ttu-id="11c76-154">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="11c76-154">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="11c76-155">Word Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="11c76-155">Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>

    - <span data-ttu-id="11c76-156">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="11c76-156">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="11c76-157">No menu **Página Inicial** do Word, selecione **Mostrar Painel de Tarefas**.</span><span class="sxs-lookup"><span data-stu-id="11c76-157">On the **Home** menu of Word, select **Show Taskpane**.</span></span>

6. <span data-ttu-id="11c76-158">No painel de tarefas, escolha **Inserir Parágrafo**.</span><span class="sxs-lookup"><span data-stu-id="11c76-158">In the task pane, choose **Insert Paragraph**.</span></span>

7. <span data-ttu-id="11c76-159">Faça uma alteração no parágrafo.</span><span class="sxs-lookup"><span data-stu-id="11c76-159">Make a change in the paragraph.</span></span>

8. <span data-ttu-id="11c76-160">Escolha novamente **Inserir Parágrafo**.</span><span class="sxs-lookup"><span data-stu-id="11c76-160">Choose **Insert Paragraph** again.</span></span> <span data-ttu-id="11c76-161">O novo parágrafo está acima do anterior porque o método `insertParagraph` está inserido no início do corpo do documento.</span><span class="sxs-lookup"><span data-stu-id="11c76-161">Note that the new paragraph is above the previous one because the `insertParagraph` method is inserting at the "start" of the document's body.</span></span>

    ![Tutorial do Word: Inserir Parágrafo](../images/word-tutorial-insert-paragraph.png)

## <a name="format-text"></a><span data-ttu-id="11c76-163">Formatar texto</span><span class="sxs-lookup"><span data-stu-id="11c76-163">Format text</span></span>

<span data-ttu-id="11c76-164">Nesta etapa do tutorial, você aplicará um estilo interno ao texto, aplicará um estilo personalizado ao texto e alterará a fonte do texto.</span><span class="sxs-lookup"><span data-stu-id="11c76-164">In this step of the tutorial, you'll apply a built-in style to text, apply a custom style to text, and change the font of text.</span></span>

### <a name="apply-a-built-in-style-to-text"></a><span data-ttu-id="11c76-165">Aplicar um estilo interno ao texto</span><span class="sxs-lookup"><span data-stu-id="11c76-165">Apply a built-in style to text</span></span>

1. <span data-ttu-id="11c76-166">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="11c76-166">Open the project in your code editor.</span></span> 

2. <span data-ttu-id="11c76-167">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="11c76-167">Open the file index.html.</span></span>

3. <span data-ttu-id="11c76-168">Abaixo do `div`, que contém o botão `insert-paragraph`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-168">Just below the `div` that contains the `insert-paragraph` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. <span data-ttu-id="11c76-169">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="11c76-169">Open the app.js file.</span></span>

5. <span data-ttu-id="11c76-170">Logo abaixo da linha que atribui um identificador de clique ao botão `insert-paragraph`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="11c76-170">Just below the line that assigns a click handler to the `insert-paragraph` button, add the following code:</span></span>

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. <span data-ttu-id="11c76-171">Logo abaixo da função `insertParagraph`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-171">Just below the `insertParagraph` function, add the following function:</span></span>

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

7. <span data-ttu-id="11c76-172">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="11c76-172">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="11c76-173">O código aplica um estilo a um parágrafo, mas também é possível aplicar estilos em intervalos de texto.</span><span class="sxs-lookup"><span data-stu-id="11c76-173">Note that the code applies a style to a paragraph, but styles can also be applied to ranges of text.</span></span>

    ```js
    var firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

### <a name="apply-a-custom-style-to-text"></a><span data-ttu-id="11c76-174">Aplicar um estilo personalizado ao texto</span><span class="sxs-lookup"><span data-stu-id="11c76-174">Apply a custom style to text</span></span>

1. <span data-ttu-id="11c76-175">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="11c76-175">Open the file index.html.</span></span>

2. <span data-ttu-id="11c76-176">Abaixo do `div` que contém o botão `apply-style`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-176">Below the `div` that contains the `apply-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. <span data-ttu-id="11c76-177">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="11c76-177">Open the app.js file.</span></span>

4. <span data-ttu-id="11c76-178">Abaixo da linha que atribui um identificador de clique ao botão `apply-style`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="11c76-178">Below the line that assigns a click handler to the `apply-style` button, add the following code:</span></span>

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. <span data-ttu-id="11c76-179">Abaixo da função `applyStyle`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-179">Below the `applyStyle` function, add the following function:</span></span>

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

6. <span data-ttu-id="11c76-180">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="11c76-180">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="11c76-181">O código aplica um estilo personalizado que ainda não existe.</span><span class="sxs-lookup"><span data-stu-id="11c76-181">Note that the code applies a custom style that does not exist yet.</span></span> <span data-ttu-id="11c76-182">Você criará um estilo com o nome **MyCustomStyle** na etapa [Testar o suplemento](#test-the-add-in).</span><span class="sxs-lookup"><span data-stu-id="11c76-182">You'll create a style with the name **MyCustomStyle** in the [Test the add-in](#test-the-add-in) step.</span></span>

    ```js
    var lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

### <a name="change-the-font-of-text"></a><span data-ttu-id="11c76-183">Alterar a fonte do texto</span><span class="sxs-lookup"><span data-stu-id="11c76-183">Change the font of text</span></span>

1. <span data-ttu-id="11c76-184">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="11c76-184">Open the file index.html.</span></span>

2. <span data-ttu-id="11c76-185">Abaixo do `div` que contém o botão `apply-custom-style`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-185">Below the `div` that contains the `apply-custom-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. <span data-ttu-id="11c76-186">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="11c76-186">Open the app.js file.</span></span>

4. <span data-ttu-id="11c76-187">Abaixo da linha que atribui um identificador de clique ao botão `apply-custom-style`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="11c76-187">Below the line that assigns a click handler to the `apply-custom-style` button, add the following code:</span></span>

    ```js
    $('#change-font').click(changeFont);
    ```

5. <span data-ttu-id="11c76-188">Abaixo da função `applyCustomStyle`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-188">Below the `applyCustomStyle` function, add the following function:</span></span>

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

6. <span data-ttu-id="11c76-189">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="11c76-189">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="11c76-190">O código recebe uma referência para o segundo parágrafo usando o método `ParagraphCollection.getFirst` encadeado para o método `Paragraph.getNext`.</span><span class="sxs-lookup"><span data-stu-id="11c76-190">Note that the code gets a reference to the second paragraph by using the `ParagraphCollection.getFirst` method chained to the `Paragraph.getNext` method.</span></span>

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

### <a name="test-the-add-in"></a><span data-ttu-id="11c76-191">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="11c76-191">Test the add-in</span></span>

1. <span data-ttu-id="11c76-192">Na janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estão abertos, digite Ctrl + C duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="11c76-192">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="11c76-193">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="11c76-193">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="11c76-194">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="11c76-194">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="11c76-195">Para fazer isso, interrompa o processo do servidor para que o prompt apareça e você possa inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="11c76-195">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="11c76-196">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="11c76-196">After the build, you restart the server.</span></span> <span data-ttu-id="11c76-197">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="11c76-197">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="11c76-198">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.</span><span class="sxs-lookup"><span data-stu-id="11c76-198">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="11c76-199">Execute o comando `npm start` para iniciar um servidor Web em um localhost.</span><span class="sxs-lookup"><span data-stu-id="11c76-199">Run the command `npm start` to start a web server running on localhost.</span></span>   

4. <span data-ttu-id="11c76-200">Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="11c76-200">Reload the task pane by closing it, and then on the **Home** menu select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="11c76-201">Verifique se há pelo menos três parágrafos no documento.</span><span class="sxs-lookup"><span data-stu-id="11c76-201">Be sure there are at least three paragraphs in the document.</span></span> <span data-ttu-id="11c76-202">É possível escolher **Inserir Parágrafo** três vezes.</span><span class="sxs-lookup"><span data-stu-id="11c76-202">You can choose **Insert Paragraph** three times.</span></span> <span data-ttu-id="11c76-203">*Verifique com atenção se não há um parágrafo em branco no final do documento. Se houver, exclua-o.*</span><span class="sxs-lookup"><span data-stu-id="11c76-203">*Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.*</span></span>

6. <span data-ttu-id="11c76-204">No Word, crie um estilo personalizado chamado "MyCustomStyle".</span><span class="sxs-lookup"><span data-stu-id="11c76-204">In Word, create a custom style named "MyCustomStyle".</span></span> <span data-ttu-id="11c76-205">Pode ter a formatação que você quiser.</span><span class="sxs-lookup"><span data-stu-id="11c76-205">It can have any formatting that you want.</span></span>

7. <span data-ttu-id="11c76-206">Escolha o botão **Aplicar Estilo**.</span><span class="sxs-lookup"><span data-stu-id="11c76-206">Choose the **Apply Style** button.</span></span> <span data-ttu-id="11c76-207">O primeiro parágrafo receberá o estilo interno **Referência Intensa**.</span><span class="sxs-lookup"><span data-stu-id="11c76-207">The first paragraph will be styled with the built-in style **Intense Reference**.</span></span>

8. <span data-ttu-id="11c76-208">Escolha o botão **Aplicar Estilo Personalizado**.</span><span class="sxs-lookup"><span data-stu-id="11c76-208">Choose the **Apply Custom Style** button.</span></span> <span data-ttu-id="11c76-209">O último parágrafo receberá seu estilo personalizado.</span><span class="sxs-lookup"><span data-stu-id="11c76-209">The last paragraph will be styled with your custom style.</span></span> <span data-ttu-id="11c76-210">(Se parecer que nada acontece, talvez o último parágrafo esteja em branco.</span><span class="sxs-lookup"><span data-stu-id="11c76-210">(If nothing seems to happen, the last paragraph might be blank.</span></span> <span data-ttu-id="11c76-211">Se estiver, adicione um texto a ele).</span><span class="sxs-lookup"><span data-stu-id="11c76-211">If so, add some text to it.)</span></span>

9. <span data-ttu-id="11c76-212">Escolha o botão **Alterar Fonte**.</span><span class="sxs-lookup"><span data-stu-id="11c76-212">Choose the **Change Font** button.</span></span> <span data-ttu-id="11c76-213">A fonte do segundo parágrafo muda para 18 pt, negrito, Courier New.</span><span class="sxs-lookup"><span data-stu-id="11c76-213">The font of the second paragraph changes to 18 pt., bold, Courier New.</span></span>

    ![Tutorial do Word: Aplicar estilos e fonte](../images/word-tutorial-apply-styles-and-font.png)

## <a name="replace-text-and-insert-text"></a><span data-ttu-id="11c76-215">Substituir texto e inserir texto</span><span class="sxs-lookup"><span data-stu-id="11c76-215">Replace text and insert text in various locations</span></span>

<span data-ttu-id="11c76-216">Nesta etapa o tutorial, você adicionará texto dentro e fora dos intervalos de texto selecionados, e substituirá o texto de um intervalo selecionado.</span><span class="sxs-lookup"><span data-stu-id="11c76-216">In this step of the tutorial, you'll add text inside and outside of selected ranges of text, and replace the text of a selected range.</span></span>

### <a name="add-text-inside-a-range"></a><span data-ttu-id="11c76-217">Adicionar texto dentro de um intervalo</span><span class="sxs-lookup"><span data-stu-id="11c76-217">Add text inside a range</span></span>

1. <span data-ttu-id="11c76-218">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="11c76-218">Open the project in your code editor.</span></span>

2. <span data-ttu-id="11c76-219">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="11c76-219">Open the file index.html.</span></span>

3. <span data-ttu-id="11c76-220">Abaixo do `div` que contém o botão `change-font`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-220">Below the `div` that contains the `change-font` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button>
    </div>
    ```

4. <span data-ttu-id="11c76-221">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="11c76-221">Open the app.js file.</span></span>

5. <span data-ttu-id="11c76-222">Abaixo da linha que atribui um identificador de clique ao botão `change-font`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="11c76-222">Below the line that assigns a click handler to the `change-font` button, add the following code:</span></span>

    ```js
    $('#insert-text-into-range').click(insertTextIntoRange);
    ```

6. <span data-ttu-id="11c76-223">Abaixo da função `changeFont`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-223">Below the `changeFont` function, add the following function:</span></span>

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

7. <span data-ttu-id="11c76-p120">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="11c76-p120">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="11c76-226">o método serve para inserir a abreviação ["(C2R)"] no final do Intervalo cujo texto é "Clique para Executar".</span><span class="sxs-lookup"><span data-stu-id="11c76-226">The method is intended to insert the abbreviation ["(C2R)"] into the end of the Range whose text is "Click-to-Run".</span></span> <span data-ttu-id="11c76-227">Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.</span><span class="sxs-lookup"><span data-stu-id="11c76-227">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="11c76-228">O primeiro parâmetro do método `Range.insertText` é a cadeia de caracteres a ser inserida no objeto `Range`.</span><span class="sxs-lookup"><span data-stu-id="11c76-228">The first parameter of the `Range.insertText` method is the string to insert into the `Range` object.</span></span>

   - <span data-ttu-id="11c76-229">O segundo parâmetro especifica onde no intervalo, o texto adicional deve ser inserido.</span><span class="sxs-lookup"><span data-stu-id="11c76-229">The second parameter specifies where in the range the additional text should be inserted.</span></span> <span data-ttu-id="11c76-230">Além de "Fim", as outras opções possíveis são "Início", "Antes", "Depois" e "Substituir".</span><span class="sxs-lookup"><span data-stu-id="11c76-230">Besides "End", the other possible options are "Start", "Before", "After", and "Replace".</span></span> 

   - <span data-ttu-id="11c76-231">A diferença entre "Fim" e "Depois" é que "Fim" insere o novo texto dentro o final do intervalo existente, mas "Depois" cria um novo intervalo com a cadeia de caracteres e insere o novo intervalo após o intervalo existente.</span><span class="sxs-lookup"><span data-stu-id="11c76-231">The difference between "End" and "After" is that "End" inserts the new text inside the end of the existing range, but "After" creates a new range with the string and inserts the new range after the existing range.</span></span> <span data-ttu-id="11c76-232">Da mesma forma, "Início" insere o texto dentro do início do intervalo existente, e "Antes" insere um novo intervalo.</span><span class="sxs-lookup"><span data-stu-id="11c76-232">Similarly, "Start" inserts text inside the beginning of the existing range and "Before" inserts a new range.</span></span> <span data-ttu-id="11c76-233">"Substituir" substitui o texto do intervalo existente pela cadeia de caracteres do primeiro parâmetro.</span><span class="sxs-lookup"><span data-stu-id="11c76-233">"Replace" replaces the text of the existing range with the string in the first parameter.</span></span>

   - <span data-ttu-id="11c76-234">Você viu em um estágio anterior do tutorial que os métodos insert\* do objeto de corpo não têm as opções "Antes" e "Depois".</span><span class="sxs-lookup"><span data-stu-id="11c76-234">You saw in an earlier stage of the tutorial that the insert\* methods of the body object do not have the "Before" and "After" options.</span></span> <span data-ttu-id="11c76-235">Isso ocorre porque não é possível colocar o conteúdo fora do corpo do documento.</span><span class="sxs-lookup"><span data-stu-id="11c76-235">This is because you can't put content outside of the document's body.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ```

8. <span data-ttu-id="11c76-236">Vamos deixar `TODO2` de lado até a próxima seção.</span><span class="sxs-lookup"><span data-stu-id="11c76-236">We'll skip over `TODO2` until the next section.</span></span> <span data-ttu-id="11c76-237">Substitua `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="11c76-237">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="11c76-238">Esse código é semelhante ao código que você criou no primeiro estágio do tutorial, exceto que, agora, você está inserindo um novo parágrafo no final do documento, em vez de no início.</span><span class="sxs-lookup"><span data-stu-id="11c76-238">This code is similar to the code you created in the first stage of the tutorial, except that now you are inserting a new paragraph at the end of the document instead of at the start.</span></span> <span data-ttu-id="11c76-239">Este novo parágrafo demonstrará que o novo texto agora faz parte do intervalo original.</span><span class="sxs-lookup"><span data-stu-id="11c76-239">This new paragraph will demonstrate that the new text is now part of the original range.</span></span>

    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text, "End");
    ```

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="11c76-240">Adicione código para buscar propriedades do documento em objetos de script do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="11c76-240">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="11c76-241">Em todas as funções anteriores desta série de tutoriais, você colocou em fila comandos para *gravar* no documento do Office.</span><span class="sxs-lookup"><span data-stu-id="11c76-241">In all the previous functions in this series of tutorials, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="11c76-242">Cada função terminou com uma chamada para o método `context.sync()`, que envia os comandos em fila para o documento a ser executado.</span><span class="sxs-lookup"><span data-stu-id="11c76-242">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="11c76-243">Entretanto, o código adicionado na última etapa chama a propriedade `originalRange.text` e essa é uma grande diferença das funções anteriores que você escreveu, pois o objeto `originalRange` é apenas um objeto de proxy que existe no script do seu painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="11c76-243">But the code you added in the last step calls the `originalRange.text` property, and this is a significant difference from the earlier functions you wrote, because the `originalRange` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="11c76-244">Ele não sabe qual é o texto real do intervalo no documento, portanto, sua propriedade `text` não pode ter um valor real.</span><span class="sxs-lookup"><span data-stu-id="11c76-244">It doesn't know what the actual text of the range in the document is, so its `text` property can't have a real value.</span></span> <span data-ttu-id="11c76-245">Primeiro, é necessário buscar o valor de texto do intervalo no documento e usá-lo para definir o valor de `originalRange.text`.</span><span class="sxs-lookup"><span data-stu-id="11c76-245">It is necessary to first fetch the text value of the range from the document and use it to set the value of `originalRange.text`.</span></span> <span data-ttu-id="11c76-246">Somente então será possível chamar `originalRange.text` sem causar uma exceção.</span><span class="sxs-lookup"><span data-stu-id="11c76-246">Only then can `originalRange.text` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="11c76-247">Esse processo de busca tem três etapas:</span><span class="sxs-lookup"><span data-stu-id="11c76-247">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="11c76-248">Coloque em fila um comando para carregar (ou seja, fetch) as propriedades que seu código precisa ler.</span><span class="sxs-lookup"><span data-stu-id="11c76-248">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="11c76-249">Chame o método `sync` do objeto de contexto para enviar o comando em fila para o documento para execução e retornar as informações solicitadas.</span><span class="sxs-lookup"><span data-stu-id="11c76-249">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="11c76-250">Como o método `sync` é assíncrono, certifique-se de que ele tenha sido concluído antes que o código chame as propriedades que foram buscadas.</span><span class="sxs-lookup"><span data-stu-id="11c76-250">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="11c76-251">Estas etapas devem ser concluídas sempre que seu código precisar *ler* informações do documento do Office.</span><span class="sxs-lookup"><span data-stu-id="11c76-251">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="11c76-252">Substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="11c76-252">Replace `TODO2` with the following code.</span></span>
  
    ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO4: Move the doc.body.insertParagraph line here.

            }
        )
            // TODO5: Move the final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has
            //        been queued.
    ```

2. <span data-ttu-id="11c76-p127">Você não pode ter duas instruções `return` no mesmo caminho de código sem ramificações, portanto, exclua a linha final `return context.sync();` no final de `Word.run`. Você adicionará um novo final `context.sync` posteriormente neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="11c76-p127">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Word.run`. You'll add a new final `context.sync` later in this tutorial.</span></span>

3. <span data-ttu-id="11c76-255">Recorte a linha `doc.body.insertParagraph` e cole no lugar de `TODO4`.</span><span class="sxs-lookup"><span data-stu-id="11c76-255">Cut the `doc.body.insertParagraph` line and paste in place of `TODO4`.</span></span>

4. <span data-ttu-id="11c76-p128">Substitua `TODO5` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="11c76-p128">Replace `TODO5` with the following code. Note:</span></span>

   - <span data-ttu-id="11c76-258">Passar o método `sync` para uma função `then` garante que ele não seja executado até que a lógica `insertParagraph` tenha sido enfileirada.</span><span class="sxs-lookup"><span data-stu-id="11c76-258">Passing the `sync` method to a `then` function ensures that it does not run until the `insertParagraph` logic has been queued.</span></span>

   - <span data-ttu-id="11c76-259">O método `then` invoca qualquer função que é passada para ele e não é recomendável que `sync` seja chamado duas vezes, portanto, remova os "()" do fim de context.sync.</span><span class="sxs-lookup"><span data-stu-id="11c76-259">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of context.sync.</span></span>

    ```js
    .then(context.sync);
    ```

<span data-ttu-id="11c76-260">Quando terminar, a função inteira deve se parecer com o seguinte:</span><span class="sxs-lookup"><span data-stu-id="11c76-260">When you are done, the entire function should look like the following:</span></span>

```js
function insertTextIntoRange() {
    Word.run(function (context) {

        var doc = context.document;
        var originalRange = doc.getSelection();
        originalRange.insertText(" (C2R)", "End");

        originalRange.load("text");
        return context.sync()
            .then(function() {
                        doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                                                "End");
                }
            )
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

### <a name="add-text-between-ranges"></a><span data-ttu-id="11c76-261">Adicionar texto entre intervalos</span><span class="sxs-lookup"><span data-stu-id="11c76-261">Add text between ranges</span></span>

1. <span data-ttu-id="11c76-262">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="11c76-262">Open the file index.html.</span></span>

2. <span data-ttu-id="11c76-263">Abaixo do `div` que contém o botão `insert-text-into-range`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-263">Below the `div` that contains the `insert-text-into-range` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button>
    </div>
    ```

3. <span data-ttu-id="11c76-264">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="11c76-264">Open the app.js file.</span></span>

4. <span data-ttu-id="11c76-265">Abaixo da linha que atribui um identificador de clique ao botão `insert-text-into-range`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="11c76-265">Below the line that assigns a click handler to the `insert-text-into-range` button, add the following code:</span></span>

    ```js
    $('#insert-text-outside-range').click(insertTextBeforeRange);
    ```

5. <span data-ttu-id="11c76-266">Abaixo da função `insertTextIntoRange`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-266">Below the `insertTextIntoRange` function, add the following function:</span></span>

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

6. <span data-ttu-id="11c76-p129">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="11c76-p129">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="11c76-269">O método serve para adicionar um intervalo cujo texto seja "Office 2019", antes do intervalo com o texto "Office 365".</span><span class="sxs-lookup"><span data-stu-id="11c76-269">The method is intended to add a range whose text is "Office 2019, " before the range with text "Office 365".</span></span> <span data-ttu-id="11c76-270">Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.</span><span class="sxs-lookup"><span data-stu-id="11c76-270">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="11c76-271">O primeiro parâmetro do método `Range.insertText` é a cadeia de caracteres a ser adicionada.</span><span class="sxs-lookup"><span data-stu-id="11c76-271">The first parameter of the `Range.insertText` method is the string to add.</span></span>

   - <span data-ttu-id="11c76-272">O segundo parâmetro especifica onde no intervalo, o texto adicional deve ser inserido.</span><span class="sxs-lookup"><span data-stu-id="11c76-272">The second parameter specifies where in the range the additional text should be inserted.</span></span> <span data-ttu-id="11c76-273">Para ter mais detalhes sobre as opções de local, confira a discussão anterior sobre a função `insertTextIntoRange`.</span><span class="sxs-lookup"><span data-stu-id="11c76-273">For more details about the location options, see the previous discussion of the `insertTextIntoRange` function.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ```

7. <span data-ttu-id="11c76-274">Substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="11c76-274">Replace `TODO2` with the following code.</span></span>

     ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO3: Queue commands to insert the original range as a
                //        paragraph at the end of the document.

                }
            )

            // TODO4: Make a final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has
            //        been queued.
    ```

8. <span data-ttu-id="11c76-275">Substitua `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="11c76-275">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="11c76-276">Este novo parágrafo demonstrará que o novo texto ***não*** faz parte do intervalo original selecionado.</span><span class="sxs-lookup"><span data-stu-id="11c76-276">This new paragraph will demonstrate the fact that the new text is ***not*** part of the original selected range.</span></span> <span data-ttu-id="11c76-277">O intervalo original ainda contém o texto que tinha quando foi selecionado.</span><span class="sxs-lookup"><span data-stu-id="11c76-277">The original range still has only the text it had when it was selected.</span></span>

    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                             "End");
    ```

9. <span data-ttu-id="11c76-278">Substitua `TODO4` pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-278">Replace `TODO4` with the following code:</span></span>

    ```js
    .then(context.sync);
    ```

### <a name="replace-the-text-of-a-range"></a><span data-ttu-id="11c76-279">Substitua o texto de um intervalo.</span><span class="sxs-lookup"><span data-stu-id="11c76-279">Replace the text of a range</span></span>

1. <span data-ttu-id="11c76-280">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="11c76-280">Open the file index.html.</span></span>

2. <span data-ttu-id="11c76-281">Abaixo do `div` que contém o botão `insert-text-outside-range`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-281">Below the `div` that contains the `insert-text-outside-range` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="replace-text">Change Quantity Term</button>
    </div>
    ```

3. <span data-ttu-id="11c76-282">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="11c76-282">Open the app.js file.</span></span>

4. <span data-ttu-id="11c76-283">Abaixo da linha que atribui um identificador de clique ao botão `insert-text-outside-range`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="11c76-283">Below the line that assigns a click handler to the `insert-text-outside-range` button, add the following code:</span></span>

    ```js
    $('#replace-text').click(replaceText);
    ```

5. <span data-ttu-id="11c76-284">Abaixo da função `insertTextBeforeRange`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-284">Below the `insertTextBeforeRange` function, add the following function:</span></span>

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

6. <span data-ttu-id="11c76-285">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="11c76-285">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="11c76-286">O método serve para substituir a cadeia de caracteres "várias" pela cadeia "muitos".</span><span class="sxs-lookup"><span data-stu-id="11c76-286">Note that the method is intended to replace the string "several" with the string "many".</span></span> <span data-ttu-id="11c76-287">Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.</span><span class="sxs-lookup"><span data-stu-id="11c76-287">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="11c76-288">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="11c76-288">Test the add-in</span></span>

1. <span data-ttu-id="11c76-289">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl + C duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="11c76-289">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="11c76-290">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="11c76-290">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="11c76-291">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="11c76-291">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="11c76-292">Para fazer isso, interrompa o processo do servidor para que o prompt apareça e você possa inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="11c76-292">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="11c76-293">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="11c76-293">After the build, restart the server.</span></span> <span data-ttu-id="11c76-294">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="11c76-294">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="11c76-295">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.</span><span class="sxs-lookup"><span data-stu-id="11c76-295">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="11c76-296">Execute o comando `npm start` para iniciar um servidor Web em um localhost.</span><span class="sxs-lookup"><span data-stu-id="11c76-296">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="11c76-297">Feche o painel de tarefas para recarregá-lo e, no menu **Página Inicial**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="11c76-297">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="11c76-298">No painel de tarefas, escolha **Inserir Parágrafo** para garantir que haja um parágrafo no início do documento.</span><span class="sxs-lookup"><span data-stu-id="11c76-298">In the task pane, choose **Insert Paragraph** to ensure that there is a paragraph at the start of the document.</span></span>

6. <span data-ttu-id="11c76-299">Selecione um texto.</span><span class="sxs-lookup"><span data-stu-id="11c76-299">Select some text.</span></span> <span data-ttu-id="11c76-300">Selecionar a frase "Clique para Executar" fará mais sentido.</span><span class="sxs-lookup"><span data-stu-id="11c76-300">Selecting the phrase "Click-to-Run" will make the most sense.</span></span> <span data-ttu-id="11c76-301">*Tenha cuidado para não incluir o espaço anterior ou seguinte na seleção.*</span><span class="sxs-lookup"><span data-stu-id="11c76-301">*Be careful not to include the preceding or following space in the selection.*</span></span>

7. <span data-ttu-id="11c76-302">Escolha o botão **Inserir Abreviação**.</span><span class="sxs-lookup"><span data-stu-id="11c76-302">Choose the **Insert Abbreviation** button.</span></span> <span data-ttu-id="11c76-303">"(C2R)" é adicionado.</span><span class="sxs-lookup"><span data-stu-id="11c76-303">Note that " (C2R)" is added.</span></span> <span data-ttu-id="11c76-304">Na parte inferior do documento, um novo parágrafo é adicionado com o texto inteiro expandido porque a nova cadeia de caracteres foi adicionada ao intervalo existente.</span><span class="sxs-lookup"><span data-stu-id="11c76-304">Note also that at the bottom of the document a new paragraph is added with the entire expanded text because the new string was added to the existing range.</span></span>

8. <span data-ttu-id="11c76-305">Selecione um texto.</span><span class="sxs-lookup"><span data-stu-id="11c76-305">Select some text.</span></span> <span data-ttu-id="11c76-306">Selecionar a frase "Office 365" fará mais sentido.</span><span class="sxs-lookup"><span data-stu-id="11c76-306">Selecting the phrase "Office 365" will make the most sense.</span></span> <span data-ttu-id="11c76-307">*Tenha cuidado para não incluir o espaço anterior ou seguinte na seleção.*</span><span class="sxs-lookup"><span data-stu-id="11c76-307">*Be careful not to include the preceding or following space in the selection.*</span></span>

9. <span data-ttu-id="11c76-308">Escolha o botão **Adicionar Informações de Versão**.</span><span class="sxs-lookup"><span data-stu-id="11c76-308">Choose the **Add Version Info** button.</span></span> <span data-ttu-id="11c76-309">"Office 2019" está inserido entre "Office 2016" e "Office 365".</span><span class="sxs-lookup"><span data-stu-id="11c76-309">Note that "Office 2019, " is inserted between "Office 2016" and "Office 365".</span></span> <span data-ttu-id="11c76-310">Na parte inferior do documento um novo parágrafo foi adicionado, mas ele contém apenas o texto selecionado originalmente porque a nova cadeia de caracteres tornou-se um intervalo novo, em vez de ser adicionada ao intervalo original.</span><span class="sxs-lookup"><span data-stu-id="11c76-310">Note also that at the bottom of the document a new paragraph is added but it contains only the originally selected text because the new string became a new range rather than being added to the original range.</span></span>

10. <span data-ttu-id="11c76-311">Selecione um texto.</span><span class="sxs-lookup"><span data-stu-id="11c76-311">Select some text.</span></span> <span data-ttu-id="11c76-312">Selecionar a palavra "vários" fará mais sentido.</span><span class="sxs-lookup"><span data-stu-id="11c76-312">Selecting the word "several" will make the most sense.</span></span> <span data-ttu-id="11c76-313">*Tenha cuidado para não incluir o espaço anterior ou seguinte na seleção.*</span><span class="sxs-lookup"><span data-stu-id="11c76-313">*Be careful not to include the preceding or following space in the selection.*</span></span>

11. <span data-ttu-id="11c76-314">Escolha o botão **Alterar Termo de Quantidade**.</span><span class="sxs-lookup"><span data-stu-id="11c76-314">Choose the **Change Quantity Term** button.</span></span> <span data-ttu-id="11c76-315">"muitos" substitui o texto selecionado.</span><span class="sxs-lookup"><span data-stu-id="11c76-315">Note that "many" replaces the selected text.</span></span>

    ![Tutorial do Word: texto adicionado e substituído](../images/word-tutorial-text-replace.png)

## <a name="insert-images-html-and-tables"></a><span data-ttu-id="11c76-317">Inserir imagens, HTML e tabelas</span><span class="sxs-lookup"><span data-stu-id="11c76-317">Insert images, HTML, and tables</span></span>

<span data-ttu-id="11c76-318">Nesta etapa do tutorial, você aprenderá a inserir imagens, HTML e tabelas no documento.</span><span class="sxs-lookup"><span data-stu-id="11c76-318">In this step of the tutorial, you'll learn how to insert images, HTML, and tables into the document.</span></span>

### <a name="insert-an-image"></a><span data-ttu-id="11c76-319">Inserir uma imagem</span><span class="sxs-lookup"><span data-stu-id="11c76-319">Insert an image</span></span>

1. <span data-ttu-id="11c76-320">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="11c76-320">Open the project in your code editor.</span></span>

2. <span data-ttu-id="11c76-321">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="11c76-321">Open the file index.html.</span></span>

3. <span data-ttu-id="11c76-322">Abaixo do `div` que contém o botão `replace-text`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-322">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-image">Insert Image</button>
    </div>
    ```

4. <span data-ttu-id="11c76-323">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="11c76-323">Open the app.js file.</span></span>

5. <span data-ttu-id="11c76-324">Na parte superior do arquivo, logo abaixo da linha use-strict, adicione a seguinte linha.</span><span class="sxs-lookup"><span data-stu-id="11c76-324">Near the top of the file, just below the use-strict line, add the following line.</span></span> <span data-ttu-id="11c76-325">Essa linha importa uma variável de outro arquivo.</span><span class="sxs-lookup"><span data-stu-id="11c76-325">This line imports a variable from another file.</span></span> <span data-ttu-id="11c76-326">A variável é uma cadeia de caracteres base 64 que codifica uma imagem.</span><span class="sxs-lookup"><span data-stu-id="11c76-326">The variable is a base 64 string that encodes an image.</span></span> <span data-ttu-id="11c76-327">Para ver a cadeia de caracteres codificada, abra o arquivo base64Image.js na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="11c76-327">To see the encoded string, open the base64Image.js file in the root of the project.</span></span>

    ```js
    import { base64Image } from "./base64Image";
    ```

6. <span data-ttu-id="11c76-328">Abaixo da linha que atribui um identificador de clique ao botão `replace-text`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="11c76-328">Below the line that assigns a click handler to the `replace-text` button, add the following code:</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

7. <span data-ttu-id="11c76-329">Abaixo da função `replaceText`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-329">Below the `replaceText` function, add the following function:</span></span>

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

8. <span data-ttu-id="11c76-330">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="11c76-330">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="11c76-331">Esta linha insere a imagem codificada em base 64 no final do documento.</span><span class="sxs-lookup"><span data-stu-id="11c76-331">Note that this line inserts the base 64 encoded image at the end of the document.</span></span> <span data-ttu-id="11c76-332">(O objeto `Paragraph` também tem um método `insertInlinePictureFromBase64` e outros métodos `insert*`.</span><span class="sxs-lookup"><span data-stu-id="11c76-332">(The `Paragraph` object also has an `insertInlinePictureFromBase64` method and other `insert*` methods.</span></span> <span data-ttu-id="11c76-333">Confira a seção insertHTML a seguir para conferir um exemplo).</span><span class="sxs-lookup"><span data-stu-id="11c76-333">See the following insertHTML section for an example.)</span></span>

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

### <a name="insert-html"></a><span data-ttu-id="11c76-334">Inserir HTML</span><span class="sxs-lookup"><span data-stu-id="11c76-334">Insert HTML</span></span>

1. <span data-ttu-id="11c76-335">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="11c76-335">Open the file index.html.</span></span>

2. <span data-ttu-id="11c76-336">Abaixo do `div` que contém o botão `insert-image`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-336">Below the `div` that contains the `insert-image` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-html">Insert HTML</button>
    </div>
    ```

3. <span data-ttu-id="11c76-337">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="11c76-337">Open the app.js file.</span></span>

4. <span data-ttu-id="11c76-338">Abaixo da linha que atribui um identificador de clique ao botão `insert-image`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="11c76-338">Below the line that assigns a click handler to the `insert-image` button, add the following code:</span></span>

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. <span data-ttu-id="11c76-339">Abaixo da função `insertImage`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-339">Below the `insertImage` function, add the following function:</span></span>

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

6. <span data-ttu-id="11c76-p144">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="11c76-p144">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="11c76-342">A primeira linha adiciona um parágrafo em branco ao final do documento.</span><span class="sxs-lookup"><span data-stu-id="11c76-342">The first line adds a blank paragraph to the end of the document.</span></span> 

   - <span data-ttu-id="11c76-343">A segunda linha insere uma cadeia de caracteres de HTML no final do parágrafo; especificamente dois parágrafos, um formatado com a fonte Verdana, e o outro com estilo padrão de documento do Word.</span><span class="sxs-lookup"><span data-stu-id="11c76-343">The second line inserts a string of HTML at the end of the paragraph; specifically two paragraphs, one formatted with Verdana font, the other with the default styling of the Word document.</span></span> <span data-ttu-id="11c76-344">(Conforme mostrado anteriormente no método `insertImage`, o objeto `context.document.body` também tem os métodos `insert*`).</span><span class="sxs-lookup"><span data-stu-id="11c76-344">(As you saw in the `insertImage` method earlier, the `context.document.body` object also has the `insert*` methods.)</span></span>

    ```js
    var blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

### <a name="insert-a-table"></a><span data-ttu-id="11c76-345">Inserir uma tabela</span><span class="sxs-lookup"><span data-stu-id="11c76-345">Insert a table</span></span>

1. <span data-ttu-id="11c76-346">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="11c76-346">Open the file index.html.</span></span>

2. <span data-ttu-id="11c76-347">Abaixo do `div` que contém o botão `insert-html`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-347">Below the `div` that contains the `insert-html` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-table">Insert Table</button>
    </div>
    ```

3. <span data-ttu-id="11c76-348">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="11c76-348">Open the app.js file.</span></span>

4. <span data-ttu-id="11c76-349">Abaixo da linha que atribui um identificador de clique ao botão `insert-html`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="11c76-349">Below the line that assigns a click handler to the `insert-html` button, add the following code:</span></span>

    ```js
    $('#insert-table').click(insertTable);
    ```

5. <span data-ttu-id="11c76-350">Abaixo da função `insertHTML`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-350">Below the `insertHTML` function, add the following function:</span></span>

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

6. <span data-ttu-id="11c76-351">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="11c76-351">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="11c76-352">Essa linha usa o método `ParagraphCollection.getFirst` para obter uma referência do primeiro parágrafo e, depois, usa o método `Paragraph.getNext` para obter uma referência para o segundo parágrafo.</span><span class="sxs-lookup"><span data-stu-id="11c76-352">Note that this line uses the `ParagraphCollection.getFirst` method to get a reference ot the first paragraph and then uses the `Paragraph.getNext` method to get a reference to the second paragraph.</span></span>

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

7. <span data-ttu-id="11c76-p147">Substitua `TODO2` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="11c76-p147">Replace `TODO2` with the following code. Note:</span></span>

   - <span data-ttu-id="11c76-355">Os dois primeiros parâmetros do método `insertTable` especificam o número de linhas e colunas.</span><span class="sxs-lookup"><span data-stu-id="11c76-355">The first two parameters of the `insertTable` method specify the number of rows and columns.</span></span>

   - <span data-ttu-id="11c76-356">O terceiro parâmetro especifica onde inserir a tabela, nesse caso, depois do parágrafo.</span><span class="sxs-lookup"><span data-stu-id="11c76-356">The third parameter specifies where to insert the table, in this case after the paragraph.</span></span>

   - <span data-ttu-id="11c76-357">O quarto parâmetro é uma matriz bidimensional que define os valores das células da tabela.</span><span class="sxs-lookup"><span data-stu-id="11c76-357">The fourth parameter is a two-dimensional array that sets the values of the table cells.</span></span>

   - <span data-ttu-id="11c76-358">A tabela terá um estilo padrão simples, mas o método `insertTable` retornará um objeto `Table` com muitos membros, e alguns deles são usados para alterar o estilo de tabela.</span><span class="sxs-lookup"><span data-stu-id="11c76-358">The table will have plain default styling, but the `insertTable` method returns a `Table` object with many members, some of which are used to style the table.</span></span>

    ```js
    var tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="11c76-359">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="11c76-359">Test the add-in</span></span>

1. <span data-ttu-id="11c76-360">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl+C duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="11c76-360">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="11c76-361">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="11c76-361">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="11c76-362">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="11c76-362">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="11c76-363">Para fazer isso, interrompa o processo do servidor para que o prompt apareça e você possa inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="11c76-363">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="11c76-364">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="11c76-364">After the build, restart the server.</span></span> <span data-ttu-id="11c76-365">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="11c76-365">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="11c76-366">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.</span><span class="sxs-lookup"><span data-stu-id="11c76-366">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="11c76-367">Execute o comando `npm start` para iniciar um servidor Web em um localhost.</span><span class="sxs-lookup"><span data-stu-id="11c76-367">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="11c76-368">Feche o painel de tarefas para recarregá-lo e, no menu **Página Inicial**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="11c76-368">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="11c76-369">No painel de tarefas, escolha **Inserir Parágrafo** pelo menos três vezes para garantir que haja alguns parágrafos no documento.</span><span class="sxs-lookup"><span data-stu-id="11c76-369">In the task pane, choose **Insert Paragraph** at least three times to ensure that there are a few paragraphs in the document.</span></span>

6. <span data-ttu-id="11c76-370">Escolha o botão **Inserir Imagem**. Uma imagem é inserida no final do documento.</span><span class="sxs-lookup"><span data-stu-id="11c76-370">Choose the **Insert Image** button and note that an image is inserted at the end of the document.</span></span>

7. <span data-ttu-id="11c76-371">Escolha o botão **Inserir HTML**. Dois parágrafos são inseridos no final do documento, e o primeiro tem a fonte Verdana.</span><span class="sxs-lookup"><span data-stu-id="11c76-371">Choose the **Insert HTML** button and note that two paragraphs are inserted at the end of the document, and that the first one has Verdana font.</span></span>

8. <span data-ttu-id="11c76-372">Escolha o botão **Inserir Tabela**. Uma tabela é inserida após o segundo parágrafo.</span><span class="sxs-lookup"><span data-stu-id="11c76-372">Choose the **Insert Table** button and note that a table is inserted after the second paragraph.</span></span>

    ![Tutorial do Word: Inserir imagem, HTML e tabela](../images/word-tutorial-insert-image-html-table.png)

## <a name="create-and-update-content-controls"></a><span data-ttu-id="11c76-374">Criar e atualizar os controles de conteúdo</span><span class="sxs-lookup"><span data-stu-id="11c76-374">Create and update content controls</span></span>

<span data-ttu-id="11c76-375">Nesta etapa do tutorial, você aprenderá a criar controles de conteúdo de Rich Text no documento e, depois, como inserir e substituir conteúdo nos controles.</span><span class="sxs-lookup"><span data-stu-id="11c76-375">In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.</span></span>

> [!NOTE]
> <span data-ttu-id="11c76-376">Há vários tipos de controles de conteúdo que podem ser adicionados a um documento do Word por meio da interface do usuário. Porém, no momento, só há suporte para controles de conteúdo de Rich Text no Word.js.</span><span class="sxs-lookup"><span data-stu-id="11c76-376">There are several types of content controls that can be added to a Word document through the UI; but currently only Rich Text content controls are supported by Word.js.</span></span>
>
> <span data-ttu-id="11c76-377">Antes de começar esta etapa do tutorial, recomendamos a criação e manipulação dos controles de conteúdo de Rich Text por meio da interface do usuário do Word, para se familiarizar com os controles e suas propriedades.</span><span class="sxs-lookup"><span data-stu-id="11c76-377">Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties.</span></span> <span data-ttu-id="11c76-378">Para saber mais detalhes, confira [Criar formulários para preenchimento ou impressão no Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span><span class="sxs-lookup"><span data-stu-id="11c76-378">For details, see [Create forms that users complete or print in Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span></span>

### <a name="create-a-content-control"></a><span data-ttu-id="11c76-379">Criar um controle de conteúdo</span><span class="sxs-lookup"><span data-stu-id="11c76-379">Create a content control</span></span>

1. <span data-ttu-id="11c76-380">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="11c76-380">Open the project in your code editor.</span></span>

2. <span data-ttu-id="11c76-381">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="11c76-381">Open the file index.html.</span></span>

3. <span data-ttu-id="11c76-382">Abaixo do `div` que contém o botão `replace-text`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-382">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-content-control">Create Content Control</button>
    </div>
    ```

4. <span data-ttu-id="11c76-383">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="11c76-383">Open the app.js file.</span></span>

5. <span data-ttu-id="11c76-384">Abaixo da linha que atribui um identificador de clique ao botão `insert-table`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="11c76-384">Below the line that assigns a click handler to the `insert-table` button, add the following code:</span></span>

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. <span data-ttu-id="11c76-385">Abaixo da função `insertTable`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-385">Below the `insertTable` function, add the following function:</span></span>

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

7. <span data-ttu-id="11c76-p151">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="11c76-p151">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="11c76-388">o código tem como objetivo dispor a frase "Office 365" em um controle de conteúdo.</span><span class="sxs-lookup"><span data-stu-id="11c76-388">This code is intended to wrap the phrase "Office 365" in a content control.</span></span> <span data-ttu-id="11c76-389">Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.</span><span class="sxs-lookup"><span data-stu-id="11c76-389">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="11c76-390">A propriedade `ContentControl.title` especifica o título visível do controle de conteúdo.</span><span class="sxs-lookup"><span data-stu-id="11c76-390">The `ContentControl.title` property specifies the visible title of the content control.</span></span>

   - <span data-ttu-id="11c76-391">A propriedade `ContentControl.tag` especifica uma marca que pode ser usada para obter uma referência a um controle de conteúdo usando o método `ContentControlCollection.getByTag`, que você usará em uma função posterior.</span><span class="sxs-lookup"><span data-stu-id="11c76-391">The `ContentControl.tag` property specifies an tag that can be used to get a reference to a content control using the `ContentControlCollection.getByTag` method, which you'll use in a later function.</span></span>

   - <span data-ttu-id="11c76-392">A propriedade `ContentControl.appearance` especifica a aparência do controle.</span><span class="sxs-lookup"><span data-stu-id="11c76-392">The `ContentControl.appearance` property specifies the visual look of the control.</span></span> <span data-ttu-id="11c76-393">Usar o valor "Tags" significa que o controle será encapsulado entre marcas de abertura e fechamento, e a marca de abertura terá o título do controle de conteúdo.</span><span class="sxs-lookup"><span data-stu-id="11c76-393">Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title.</span></span> <span data-ttu-id="11c76-394">Outros valores possíveis são "BoundingBox" e "None".</span><span class="sxs-lookup"><span data-stu-id="11c76-394">Other possible values are "BoundingBox" and "None".</span></span>

   - <span data-ttu-id="11c76-395">A propriedade `ContentControl.color` especifica a cor das marcas ou da borda da caixa delimitadora.</span><span class="sxs-lookup"><span data-stu-id="11c76-395">The `ContentControl.color` property specifies the color of the tags or the border of the bounding box.</span></span>

    ```js
    var serviceNameRange = context.document.getSelection();
    var serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ```

### <a name="replace-the-content-of-the-content-control"></a><span data-ttu-id="11c76-396">Substituir o conteúdo do controle de conteúdo</span><span class="sxs-lookup"><span data-stu-id="11c76-396">Replace the content of the content control</span></span>

1. <span data-ttu-id="11c76-397">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="11c76-397">Open the file index.html.</span></span>

2. <span data-ttu-id="11c76-398">Abaixo do `div` que contém o botão `create-content-control`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-398">Below the `div` that contains the `create-content-control` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>
    </div>
    ```

3. <span data-ttu-id="11c76-399">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="11c76-399">Open the app.js file.</span></span>

4. <span data-ttu-id="11c76-400">Abaixo da linha que atribui um identificador de clique ao botão `create-content-control`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="11c76-400">Below the line that assigns a click handler to the `create-content-control` button, add the following code:</span></span>

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

5. <span data-ttu-id="11c76-401">Abaixo da função `createContentControl`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="11c76-401">Below the `createContentControl` function, add the following function:</span></span>

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

6. <span data-ttu-id="11c76-p154">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="11c76-p154">Replace `TODO1` with the following code. Note:</span></span>

    - <span data-ttu-id="11c76-404">O método `ContentControlCollection.getByTag` retorna um `ContentControlCollection` de todos os controles de conteúdo da marca especificada.</span><span class="sxs-lookup"><span data-stu-id="11c76-404">The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag.</span></span> <span data-ttu-id="11c76-405">Usamos `getFirst` para obter uma referência do controle desejado.</span><span class="sxs-lookup"><span data-stu-id="11c76-405">We use `getFirst` to get a reference to the desired control.</span></span>

    ```js
    var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="11c76-406">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="11c76-406">Test the add-in</span></span>

1. <span data-ttu-id="11c76-407">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl+C duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="11c76-407">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="11c76-408">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="11c76-408">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="11c76-409">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="11c76-409">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="11c76-410">Para fazer isso, interrompa o processo do servidor para que o prompt apareça e você possa inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="11c76-410">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="11c76-411">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="11c76-411">After the build, restart the server.</span></span> <span data-ttu-id="11c76-412">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="11c76-412">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="11c76-413">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.</span><span class="sxs-lookup"><span data-stu-id="11c76-413">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="11c76-414">Execute o comando `npm start` para iniciar um servidor Web em um localhost.</span><span class="sxs-lookup"><span data-stu-id="11c76-414">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="11c76-415">Feche o painel de tarefas para recarregá-lo e, no menu **Página Inicial**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="11c76-415">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="11c76-416">No painel de tarefas, escolha **Inserir Parágrafo** para garantir que haja um parágrafo com "Office 365" no início do documento.</span><span class="sxs-lookup"><span data-stu-id="11c76-416">In the task pane, choose **Insert Paragraph** to ensure that there is a paragraph with "Office 365" at the top of the document.</span></span>

6. <span data-ttu-id="11c76-417">Selecione a frase "Office 365" no parágrafo que você adicionou e escolha o botão **Criar Controle de Conteúdo**.</span><span class="sxs-lookup"><span data-stu-id="11c76-417">Select the phrase "Office 365" in the paragraph you just added, and then choose the **Create Content Control** button.</span></span> <span data-ttu-id="11c76-418">A frase está envolvida por marcas chamadas "Nome do Serviço".</span><span class="sxs-lookup"><span data-stu-id="11c76-418">Note that the phrase is wrapped in tags labelled "Service Name".</span></span>

7. <span data-ttu-id="11c76-419">Escolha o botão **Renomear Serviço**. O texto do controle de conteúdo muda para "Fabrikam Online Productivity Suite".</span><span class="sxs-lookup"><span data-stu-id="11c76-419">Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".</span></span>

    ![Tutorial do Word - Criar o controle de conteúdo e alterar seu texto](../images/word-tutorial-content-control.png)

## <a name="next-steps"></a><span data-ttu-id="11c76-421">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="11c76-421">Next steps</span></span>

<span data-ttu-id="11c76-422">Neste tutorial, você criou um suplemento do painel de tarefas do Word que insere e substitui texto, imagens e outro conteúdo em um documento do Word.</span><span class="sxs-lookup"><span data-stu-id="11c76-422">In this tutorial, you've created a Word task pane add-in that inserts and replaces text, images, and other content in a Word document.</span></span> <span data-ttu-id="11c76-423">Para saber mais sobre o desenvolvimento de suplementos do Word, continue no seguinte artigo:</span><span class="sxs-lookup"><span data-stu-id="11c76-423">To learn more about developing Outlook add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="11c76-424">Visão geral dos suplementos do Word</span><span class="sxs-lookup"><span data-stu-id="11c76-424">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
