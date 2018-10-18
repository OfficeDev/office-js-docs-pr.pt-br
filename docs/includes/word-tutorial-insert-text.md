<span data-ttu-id="a4f09-101">Nesta etapa do tutorial, você testará programaticamente se o suplemento oferece suporte à versão atual do Word do usuário e inserirá um parágrafo no documento.</span><span class="sxs-lookup"><span data-stu-id="a4f09-101">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Word, and then insert a paragraph in the document.</span></span>

> [!NOTE]
> <span data-ttu-id="a4f09-p101">Esta página descreve uma etapa individual de um tutorial de suplemento do Word. Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou por outro link direto, acesse a página de introdução do [tutorial de suplemento do Word](../tutorials/word-tutorial.yml) para começar pelo início.</span><span class="sxs-lookup"><span data-stu-id="a4f09-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="code-the-add-in"></a><span data-ttu-id="a4f09-104">Codificação do suplemento</span><span class="sxs-lookup"><span data-stu-id="a4f09-104">Code the add-in</span></span>

1. <span data-ttu-id="a4f09-105">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="a4f09-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="a4f09-106">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="a4f09-106">Open the file index.html.</span></span>
3. <span data-ttu-id="a4f09-107">Substitua `TODO1` pela marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="a4f09-107">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. <span data-ttu-id="a4f09-108">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="a4f09-108">Open the app.js file.</span></span>
5. <span data-ttu-id="a4f09-109">Substitua o `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="a4f09-109">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="a4f09-110">O código determina se a versão do Word do usuário suporta uma versão do Word.js que inclui todas as APIs usadas em todos os estágios deste tutorial.dae</span><span class="sxs-lookup"><span data-stu-id="a4f09-110">This code determines whether the user's version of Word supports a version of Word.js that includes all the APIs that are used in all the stages of this tutorial.</span></span> <span data-ttu-id="a4f09-111">Em um suplemento de produção, use o corpo do bloco condicional para ocultar ou desabilitar a interface do usuário que chame a APIs sem suporte.</span><span class="sxs-lookup"><span data-stu-id="a4f09-111">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="a4f09-112">Isso permitirá que o usuário ainda use as partes do suplemento às quais a versão do Word dá suporte.</span><span class="sxs-lookup"><span data-stu-id="a4f09-112">This will enable the user to still use the parts of the add-in that are supported by their version of Word.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    } 
    ```

6. <span data-ttu-id="a4f09-113">Substitua o `TODO2` pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="a4f09-113">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. <span data-ttu-id="a4f09-114">Substitua o `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="a4f09-114">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="a4f09-115">Observe o seguinte:</span><span class="sxs-lookup"><span data-stu-id="a4f09-115">Note the following:</span></span>
   - <span data-ttu-id="a4f09-116">A lógica de negócios de Word.js será adicionada à função que passar por `Word.run`.</span><span class="sxs-lookup"><span data-stu-id="a4f09-116">Your Word.js business logic will be added to the function that is passed to `Word.run`.</span></span> <span data-ttu-id="a4f09-117">Essa lógica não é executada imediatamente.</span><span class="sxs-lookup"><span data-stu-id="a4f09-117">This logic does not execute immediately.</span></span> <span data-ttu-id="a4f09-118">Em vez disso, ela é adicionada à fila de comandos pendentes.</span><span class="sxs-lookup"><span data-stu-id="a4f09-118">Instead, it is added to a queue of pending commands.</span></span>
   - <span data-ttu-id="a4f09-119">O método `context.sync` envia todos os comandos da fila para execução no Word.</span><span class="sxs-lookup"><span data-stu-id="a4f09-119">The `context.sync` method sends all queued commands to Word for execution.</span></span>
   - <span data-ttu-id="a4f09-120">O `Word.run` é seguido por um bloco `catch`.</span><span class="sxs-lookup"><span data-stu-id="a4f09-120">The `Word.run` is followed by a `catch` block.</span></span> <span data-ttu-id="a4f09-121">Essa é uma prática recomendada que você sempre deve seguir.</span><span class="sxs-lookup"><span data-stu-id="a4f09-121">This is a best practice that you should always follow.</span></span> 

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

8. <span data-ttu-id="a4f09-p106">Substitua `TODO4` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="a4f09-p106">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="a4f09-124">O primeiro parâmetro para o método `insertParagraph` é o texto para o novo parágrafo.</span><span class="sxs-lookup"><span data-stu-id="a4f09-124">The first parameter to the `insertParagraph` method is the text for the new paragraph.</span></span>
   - <span data-ttu-id="a4f09-125">O segundo parâmetro é o local dentro do corpo onde o parágrafo será inserido.</span><span class="sxs-lookup"><span data-stu-id="a4f09-125">The second parameter is the location within the body where the paragraph will be inserted.</span></span> <span data-ttu-id="a4f09-126">Outras opções para inserir parágrafo, quando o objeto pai é o corpo, são "End" e "Replace".</span><span class="sxs-lookup"><span data-stu-id="a4f09-126">Other options for insert paragraph, when the parent object is the body, are "End" and "Replace".</span></span> 

    ```js
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office Online.",
                            "Start");   
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="a4f09-127">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="a4f09-127">Test the add-in</span></span>

1. <span data-ttu-id="a4f09-128">Abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="a4f09-128">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
2. <span data-ttu-id="a4f09-129">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.</span><span class="sxs-lookup"><span data-stu-id="a4f09-129">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="a4f09-130">Execute o comando `npm start` para iniciar um servidor Web em um localhost.</span><span class="sxs-lookup"><span data-stu-id="a4f09-130">Run the command `npm start` to start a web server running on localhost.</span></span>   
4. <span data-ttu-id="a4f09-131">Realize o sideload do suplemento usando um dos métodos a seguir:</span><span class="sxs-lookup"><span data-stu-id="a4f09-131">Sideload the add-in by using one of the following methods:</span></span>
    - <span data-ttu-id="a4f09-132">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="a4f09-132">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="a4f09-133">Word Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="a4f09-133">Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="a4f09-134">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="a4f09-134">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
5. <span data-ttu-id="a4f09-135">No menu **Página Inicial** do Word, selecione **Mostrar Painel de Tarefas**.</span><span class="sxs-lookup"><span data-stu-id="a4f09-135">On the **Home** menu of Word, select **Show Taskpane**.</span></span>
6. <span data-ttu-id="a4f09-136">No painel de tarefas, escolha **Inserir Parágrafo**.</span><span class="sxs-lookup"><span data-stu-id="a4f09-136">In the taskpane, choose **Insert Paragraph**.</span></span>
7. <span data-ttu-id="a4f09-137">Faça uma alteração no parágrafo.</span><span class="sxs-lookup"><span data-stu-id="a4f09-137">Make a change in the paragraph.</span></span> 
8. <span data-ttu-id="a4f09-138">Escolha novamente **Inserir Parágrafo**.</span><span class="sxs-lookup"><span data-stu-id="a4f09-138">Choose **Insert Paragraph** again.</span></span> <span data-ttu-id="a4f09-139">O novo parágrafo está acima do anterior porque o método `insertParagraph` está inserido no "início" do corpo do documento.</span><span class="sxs-lookup"><span data-stu-id="a4f09-139">Note that the new paragraph is above the previous one because the `insertParagraph` method is inserting at the "start" of the document's body.</span></span>

    ![Tutorial do Word: Inserir Parágrafo](../images/word-tutorial-insert-paragraph.png)
