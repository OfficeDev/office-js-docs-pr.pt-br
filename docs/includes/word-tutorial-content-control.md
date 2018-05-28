<span data-ttu-id="7f573-101">Nesta etapa do tutorial, voc? aprender? a criar controles de conte?do de Rich Text no documento e, depois, como inserir e substituir conte?do nos controles.</span><span class="sxs-lookup"><span data-stu-id="7f573-101">In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.</span></span> 

> [!NOTE]
> <span data-ttu-id="7f573-p101">Esta p?gina descreve uma etapa individual de um tutorial de suplemento do Word. Se voc? chegou aqui por meio dos resultados de mecanismos de pesquisa ou por outro link direto, acesse a p?gina de introdu??o do [tutorial de suplemento do Word](../tutorials/word-tutorial.yml) para come?ar pelo in?cio.</span><span class="sxs-lookup"><span data-stu-id="7f573-p101">This page describes an individual step of a Word add-in tutorial. If you?ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

<span data-ttu-id="7f573-104">Antes de come?ar esta etapa do tutorial, recomendamos a cria??o e manipula??o dos controles de conte?do de Rich Text por meio da interface do usu?rio do Word, para se familiarizar com os controles e suas propriedades.</span><span class="sxs-lookup"><span data-stu-id="7f573-104">Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties.</span></span> <span data-ttu-id="7f573-105">Para saber mais detalhes, confira [Criar formul?rios para preenchimento ou impress?o no Word](https://support.office.com/en-us/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span><span class="sxs-lookup"><span data-stu-id="7f573-105">For details, see [Create forms that users complete or print in Word](https://support.office.com/en-us/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span></span>

> [!NOTE]
> <span data-ttu-id="7f573-106">H? v?rios tipos de controles de conte?do que podem ser adicionados a um documento do Word por meio da interface do usu?rio. Por?m, no momento, s? h? suporte para controles de conte?do de Rich Text no Word.js.</span><span class="sxs-lookup"><span data-stu-id="7f573-106">There are several types of content controls that can be added to a Word document through the UI; but currently only Rich Text content controls are supported by Word.js.</span></span>


## <a name="create-a-content-control"></a><span data-ttu-id="7f573-107">Criar um controle de conte?do</span><span class="sxs-lookup"><span data-stu-id="7f573-107">Create a content control</span></span>

1. <span data-ttu-id="7f573-108">Abra o projeto em seu editor de c?digo.</span><span class="sxs-lookup"><span data-stu-id="7f573-108">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="7f573-109">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="7f573-109">Open the file index.html.</span></span>
3. <span data-ttu-id="7f573-110">Abaixo do `div` que cont?m o bot?o `replace-text`, adicione a marca??o a seguir:</span><span class="sxs-lookup"><span data-stu-id="7f573-110">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="create-content-control">Create Content Control</button>            
    </div>
    ```

4. <span data-ttu-id="7f573-111">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="7f573-111">Open the app.js file.</span></span>

5. <span data-ttu-id="7f573-112">Abaixo da linha que atribui um identificador de clique ao bot?o `insert-table`, adicione o seguinte c?digo:</span><span class="sxs-lookup"><span data-stu-id="7f573-112">Below the line that assigns a click handler to the `insert-table` button, add the following code:</span></span>

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. <span data-ttu-id="7f573-113">Abaixo da fun??o `insertTable`, adicione a fun??o a seguir:</span><span class="sxs-lookup"><span data-stu-id="7f573-113">Below the `insertTable` function, add the following function:</span></span>

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

7. <span data-ttu-id="7f573-p103"> Substitua `TODO1` pelo c?digo a seguir. Observa??o:</span><span class="sxs-lookup"><span data-stu-id="7f573-p103">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="7f573-116">o c?digo tem como objetivo dispor a frase "Office 365" em um controle de conte?do.</span><span class="sxs-lookup"><span data-stu-id="7f573-116">This code is intended to wrap the phrase "Office 365" in a content control.</span></span> <span data-ttu-id="7f573-117">Para simplificar, ele faz uma pressuposi??o de que a cadeia de caracteres est? presente, e que o usu?rio a selecionou.</span><span class="sxs-lookup"><span data-stu-id="7f573-117">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>
   - <span data-ttu-id="7f573-118">A propriedade `ContentControl.title` especifica o t?tulo vis?vel do controle de conte?do.</span><span class="sxs-lookup"><span data-stu-id="7f573-118">The `ContentControl.title` property specifies the visible title of the content control.</span></span> 
   - <span data-ttu-id="7f573-119">A propriedade `ContentControl.tag` especifica uma marca que pode ser usada para obter uma refer?ncia a um controle de conte?do usando o m?todo `ContentControlCollection.getByTag`, que voc? usar? em uma fun??o posterior.</span><span class="sxs-lookup"><span data-stu-id="7f573-119">The `ContentControl.tag` property specifies an tag that can be used to get a reference to a content control using the `ContentControlCollection.getByTag` method, which you'll use in a later function.</span></span> 
   - <span data-ttu-id="7f573-120">A propriedade `ContentControl.appearance` especifica a apar?ncia do controle.</span><span class="sxs-lookup"><span data-stu-id="7f573-120">The `ContentControl.appearance` property specifies the visual look of the control.</span></span> <span data-ttu-id="7f573-121">Usar o valor "Tags" significa que o controle ser? encapsulado entre marcas de abertura e fechamento, e a marca de abertura ter? o t?tulo do controle de conte?do.</span><span class="sxs-lookup"><span data-stu-id="7f573-121">Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title.</span></span> <span data-ttu-id="7f573-122">Outros valores poss?veis s?o "BoundingBox" e "None".</span><span class="sxs-lookup"><span data-stu-id="7f573-122">Other possible values are "BoundingBox" and "None".</span></span>
   - <span data-ttu-id="7f573-123">A propriedade `ContentControl.color` especifica a cor das marcas ou da borda da caixa delimitadora.</span><span class="sxs-lookup"><span data-stu-id="7f573-123">The `ContentControl.color` property specifies the color of the tags or the border of the bounding box.</span></span>

    ```js
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ``` 

## <a name="replace-the-content-of-the-content-control"></a><span data-ttu-id="7f573-124">Substituir o conte?do do controle de conte?do</span><span class="sxs-lookup"><span data-stu-id="7f573-124">Replace the content of the content control</span></span>

1. <span data-ttu-id="7f573-125">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="7f573-125">Open the file index.html.</span></span>
3. <span data-ttu-id="7f573-126">Abaixo do `div` que cont?m o bot?o `create-content-control`, adicione a marca??o a seguir:</span><span class="sxs-lookup"><span data-stu-id="7f573-126">Below the `div` that contains the `create-content-control` button, add the following markup:</span></span>
    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>            
    </div>
    ```

4. <span data-ttu-id="7f573-127">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="7f573-127">Open the app.js file.</span></span>

5. <span data-ttu-id="7f573-128">Abaixo da linha que atribui um identificador de clique ao bot?o `create-content-control`, adicione o seguinte c?digo:</span><span class="sxs-lookup"><span data-stu-id="7f573-128">Below the line that assigns a click handler to the `create-content-control` button, add the following code:</span></span>

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

6. <span data-ttu-id="7f573-129">Abaixo da fun??o `createContentControl`, adicione a fun??o a seguir:</span><span class="sxs-lookup"><span data-stu-id="7f573-129">Below the `createContentControl` function, add the following function:</span></span>

    <span data-ttu-id="7f573-130">\`\`\`js    fun??o replaceContentInControl() {      Word.run(fun??o) (contexto) {</span><span class="sxs-lookup"><span data-stu-id="7f573-130">\`\`\`js    function replaceContentInControl() {      Word.run(function (context) {</span></span>
            
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
    <span data-ttu-id="7f573-131">}</span><span class="sxs-lookup"><span data-stu-id="7f573-131"></span></span>
    ``` 

7. Replace `TODO1` with the following code. 
    > [!NOTE]
    > The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag. We use `getFirst` to get a reference to the desired control.

    ```js
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="7f573-132">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="7f573-132">Test the add-in</span></span>

1. <span data-ttu-id="7f573-133">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl+C duas vezes para interromper a execu??o do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="7f573-133">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="7f573-134">Caso contr?rio, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue at? a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="7f573-134">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
     > [!NOTE]
     > <span data-ttu-id="7f573-135">Embora o servidor de sincroniza??o do navegador recarregue o suplemento no painel de tarefas sempre que voc? fizer uma altera??o em algum arquivo, incluindo o arquivo app.js, ele n?o transcompila o JavaScript, portanto, ? necess?rio repetir o comando de compila??o para que as altera??es em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="7f573-135">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="7f573-136">Para fazer isso, interrompa o processo do servidor para que o prompt apare?a e voc? possa inserir o comando de compila??o.</span><span class="sxs-lookup"><span data-stu-id="7f573-136">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="7f573-137">Ap?s a compila??o, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="7f573-137">After the build, restart the server.</span></span> <span data-ttu-id="7f573-138">As pr?ximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="7f573-138">The next few steps carry out this process.</span></span>
2. <span data-ttu-id="7f573-139">Execute o comando `npm run build` para transcompilar seu c?digo-fonte ES6 para uma vers?o anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.</span><span class="sxs-lookup"><span data-stu-id="7f573-139">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="7f573-140">Execute o comando `npm start` para iniciar um servidor Web em um localhost.</span><span class="sxs-lookup"><span data-stu-id="7f573-140">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="7f573-141">Feche o painel de tarefas para recarreg?-lo e, no menu **In?cio**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="7f573-141">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="7f573-142">No painel de tarefas, escolha **Inserir Par?grafo** para garantir que haja um par?grafo com "Office 365" no in?cio do documento.</span><span class="sxs-lookup"><span data-stu-id="7f573-142">In the taskpane, choose **Insert Paragraph** to ensure that there is a paragraph with "Office 365" at the top of the document.</span></span>
6. <span data-ttu-id="7f573-143">Selecione a frase "Office 365" no par?grafo que voc? adicionou e escolha o bot?o **Criar Controle de Conte?do**.</span><span class="sxs-lookup"><span data-stu-id="7f573-143">Select the phrase "Office 365" in the paragraph you just added, and then choose the **Create Content Control** button.</span></span> <span data-ttu-id="7f573-144">A frase est? envolvida por marcas chamadas "Nome do Servi?o".</span><span class="sxs-lookup"><span data-stu-id="7f573-144">Note that the phrase is wrapped in tags labelled "Service Name".</span></span>
7. <span data-ttu-id="7f573-145">Escolha o bot?o **Renomear Servi?o**. O texto do controle de conte?do muda para "Fabrikam Online Productivity Suite".</span><span class="sxs-lookup"><span data-stu-id="7f573-145">Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".</span></span>

    ![Tutorial do Word - Criar o controle de conte?do e alterar seu texto](../images/word-tutorial-content-control.png)
