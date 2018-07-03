<span data-ttu-id="cf350-101">Nesta etapa do tutorial, você aprenderá a criar controles de conteúdo de Rich Text no documento e, depois, como inserir e substituir conteúdo nos controles.</span><span class="sxs-lookup"><span data-stu-id="cf350-101">In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.</span></span> 

> [!NOTE]
> <span data-ttu-id="cf350-p101">Esta página descreve uma etapa individual de um tutorial de suplemento do Word. Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou por outro link direto, acesse a página de introdução do [tutorial de suplemento do Word](../tutorials/word-tutorial.yml) para começar pelo início.</span><span class="sxs-lookup"><span data-stu-id="cf350-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

<span data-ttu-id="cf350-104">Antes de começar esta etapa do tutorial, recomendamos a criação e manipulação dos controles de conteúdo de Rich Text por meio da interface do usuário do Word, para se familiarizar com os controles e suas propriedades.</span><span class="sxs-lookup"><span data-stu-id="cf350-104">Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties.</span></span> <span data-ttu-id="cf350-105">Para saber mais detalhes, confira [Criar formulários para preenchimento ou impressão no Word](https://support.office.com/en-us/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span><span class="sxs-lookup"><span data-stu-id="cf350-105">For details, see [Create forms that users complete or print in Word](https://support.office.com/en-us/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span></span>

> [!NOTE]
> <span data-ttu-id="cf350-106">Há vários tipos de controles de conteúdo que podem ser adicionados a um documento do Word por meio da interface do usuário. Porém, no momento, só há suporte para controles de conteúdo de Rich Text no Word.js.</span><span class="sxs-lookup"><span data-stu-id="cf350-106">There are several types of content controls that can be added to a Word document through the UI; but currently only Rich Text content controls are supported by Word.js.</span></span>


## <a name="create-a-content-control"></a><span data-ttu-id="cf350-107">Criar um controle de conteúdo</span><span class="sxs-lookup"><span data-stu-id="cf350-107">Create a content control</span></span>

1. <span data-ttu-id="cf350-108">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="cf350-108">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="cf350-109">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="cf350-109">Open the file index.html.</span></span>
3. <span data-ttu-id="cf350-110">Abaixo do `div` que contém o botão `replace-text`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="cf350-110">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="create-content-control">Create Content Control</button>            
    </div>
    ```

4. <span data-ttu-id="cf350-111">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="cf350-111">Open the app.js file.</span></span>

5. <span data-ttu-id="cf350-112">Abaixo da linha que atribui um identificador de clique ao botão `insert-table`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="cf350-112">Below the line that assigns a click handler to the `insert-table` button, add the following code:</span></span>

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. <span data-ttu-id="cf350-113">Abaixo da função `insertTable`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="cf350-113">Below the `insertTable` function, add the following function:</span></span>

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

7. <span data-ttu-id="cf350-p103">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="cf350-p103">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="cf350-116">o código tem como objetivo dispor a frase "Office 365" em um controle de conteúdo.</span><span class="sxs-lookup"><span data-stu-id="cf350-116">This code is intended to wrap the phrase "Office 365" in a content control.</span></span> <span data-ttu-id="cf350-117">Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.</span><span class="sxs-lookup"><span data-stu-id="cf350-117">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>
   - <span data-ttu-id="cf350-118">A propriedade `ContentControl.title` especifica o título visível do controle de conteúdo.</span><span class="sxs-lookup"><span data-stu-id="cf350-118">The `ContentControl.title` property specifies the visible title of the content control.</span></span> 
   - <span data-ttu-id="cf350-119">A propriedade `ContentControl.tag` especifica uma marca que pode ser usada para obter uma referência a um controle de conteúdo usando o método `ContentControlCollection.getByTag`, que você usará em uma função posterior.</span><span class="sxs-lookup"><span data-stu-id="cf350-119">The `ContentControl.tag` property specifies an tag that can be used to get a reference to a content control using the `ContentControlCollection.getByTag` method, which you'll use in a later function.</span></span> 
   - <span data-ttu-id="cf350-120">A propriedade `ContentControl.appearance` especifica a aparência do controle.</span><span class="sxs-lookup"><span data-stu-id="cf350-120">The `ContentControl.appearance` property specifies the visual look of the control.</span></span> <span data-ttu-id="cf350-121">Usar o valor "Tags" significa que o controle será encapsulado entre marcas de abertura e fechamento, e a marca de abertura terá o título do controle de conteúdo.</span><span class="sxs-lookup"><span data-stu-id="cf350-121">Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title.</span></span> <span data-ttu-id="cf350-122">Outros valores possíveis são "BoundingBox" e "None".</span><span class="sxs-lookup"><span data-stu-id="cf350-122">Other possible values are "BoundingBox" and "None".</span></span>
   - <span data-ttu-id="cf350-123">A propriedade `ContentControl.color` especifica a cor das marcas ou da borda da caixa delimitadora.</span><span class="sxs-lookup"><span data-stu-id="cf350-123">The `ContentControl.color` property specifies the color of the tags or the border of the bounding box.</span></span>

    ```js
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ``` 

## <a name="replace-the-content-of-the-content-control"></a><span data-ttu-id="cf350-124">Substituir o conteúdo do controle de conteúdo</span><span class="sxs-lookup"><span data-stu-id="cf350-124">Replace the content of the content control</span></span>

1. <span data-ttu-id="cf350-125">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="cf350-125">Open the file index.html.</span></span>
2. <span data-ttu-id="cf350-126">Abaixo do `div` que contém o botão `create-content-control`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="cf350-126">Below the `div` that contains the `create-content-control` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>            
    </div>
    ```

3. <span data-ttu-id="cf350-127">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="cf350-127">Open the app.js file.</span></span>

4. <span data-ttu-id="cf350-128">Abaixo da linha que atribui um identificador de clique ao botão `create-content-control`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="cf350-128">Below the line that assigns a click handler to the `create-content-control` button, add the following code:</span></span>

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

5. <span data-ttu-id="cf350-129">Abaixo da função `createContentControl`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="cf350-129">Below the `createContentControl` function, add the following function:</span></span>

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

7. <span data-ttu-id="cf350-130">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="cf350-130">Replace `TODO1` with the following code.</span></span> 
    > [!NOTE]
    > <span data-ttu-id="cf350-131">O método `ContentControlCollection.getByTag` retorna um `ContentControlCollection` de todos os controles de conteúdo da marca especificada.</span><span class="sxs-lookup"><span data-stu-id="cf350-131">The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag.</span></span> <span data-ttu-id="cf350-132">Nós usamos `getFirst` para obter uma referência do controle desejado.</span><span class="sxs-lookup"><span data-stu-id="cf350-132">We use `getFirst` to get a reference to the desired control.</span></span>

    ```js
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="cf350-133">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="cf350-133">Test the add-in</span></span>

1. <span data-ttu-id="cf350-134">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl+C duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="cf350-134">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="cf350-135">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="cf350-135">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
     > [!NOTE]
     > <span data-ttu-id="cf350-136">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="cf350-136">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="cf350-137">Para fazer isso, interrompa o processo do servidor para que o prompt apareça e você possa inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="cf350-137">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="cf350-138">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="cf350-138">After the build, restart the server.</span></span> <span data-ttu-id="cf350-139">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="cf350-139">The next few steps carry out this process.</span></span>
2. <span data-ttu-id="cf350-140">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.</span><span class="sxs-lookup"><span data-stu-id="cf350-140">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="cf350-141">Execute o comando `npm start` para iniciar um servidor Web em um localhost.</span><span class="sxs-lookup"><span data-stu-id="cf350-141">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="cf350-142">Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="cf350-142">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="cf350-143">No painel de tarefas, escolha **Inserir Parágrafo** para garantir que haja um parágrafo com "Office 365" no início do documento.</span><span class="sxs-lookup"><span data-stu-id="cf350-143">In the taskpane, choose **Insert Paragraph** to ensure that there is a paragraph with "Office 365" at the top of the document.</span></span>
6. <span data-ttu-id="cf350-144">Selecione a frase "Office 365" no parágrafo que você adicionou e escolha o botão **Criar Controle de Conteúdo**.</span><span class="sxs-lookup"><span data-stu-id="cf350-144">Select the phrase "Office 365" in the paragraph you just added, and then choose the **Create Content Control** button.</span></span> <span data-ttu-id="cf350-145">A frase está envolvida por marcas chamadas "Nome do Serviço".</span><span class="sxs-lookup"><span data-stu-id="cf350-145">Note that the phrase is wrapped in tags labelled "Service Name".</span></span>
7. <span data-ttu-id="cf350-146">Escolha o botão **Renomear Serviço**. O texto do controle de conteúdo muda para "Fabrikam Online Productivity Suite".</span><span class="sxs-lookup"><span data-stu-id="cf350-146">Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".</span></span>

    ![Tutorial do Word - Criar o controle de conteúdo e alterar seu texto](../images/word-tutorial-content-control.png)
