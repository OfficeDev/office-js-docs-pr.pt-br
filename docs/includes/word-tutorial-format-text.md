<span data-ttu-id="34fc5-101">Nesta etapa do tutorial, você mudará a fonte do texto e usará estilos internos e personalizados no texto.</span><span class="sxs-lookup"><span data-stu-id="34fc5-101">In this step of the tutorial, you'll change the font of text, and use both built-in and custom styles on the text.</span></span>

> [!NOTE]
> <span data-ttu-id="34fc5-p101">Esta página descreve uma etapa individual de um tutorial de suplemento do Word. Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou por outro link direto, acesse a página de introdução do [tutorial de suplemento do Word](../tutorials/word-tutorial.yml) para começar pelo início.</span><span class="sxs-lookup"><span data-stu-id="34fc5-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="apply-a-built-in-style-to-text"></a><span data-ttu-id="34fc5-104">Aplicar um estilo interno ao texto</span><span class="sxs-lookup"><span data-stu-id="34fc5-104">Apply a built-in style to text</span></span>

1. <span data-ttu-id="34fc5-105">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="34fc5-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="34fc5-106">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="34fc5-106">Open the file index.html.</span></span>
3. <span data-ttu-id="34fc5-107">Abaixo do `div`, que contém o botão `insert-paragraph`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="34fc5-107">Just below the `div` that contains the `insert-paragraph` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. <span data-ttu-id="34fc5-108">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="34fc5-108">Open the app.js file.</span></span>

5. <span data-ttu-id="34fc5-109">Logo abaixo da linha que atribui um identificador de clique ao botão `insert-paragraph`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="34fc5-109">Just below the line that assigns a click handler to the `insert-paragraph` button, add the following code:</span></span>

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. <span data-ttu-id="34fc5-110">Logo abaixo da função `insertParagraph`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="34fc5-110">Just below the `insertParagraph` function, add the following function:</span></span>

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

7. <span data-ttu-id="34fc5-111">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="34fc5-111">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="34fc5-112">O código aplica um estilo a um parágrafo, mas também é possível aplicar estilos em intervalos de texto.</span><span class="sxs-lookup"><span data-stu-id="34fc5-112">Note that the code applies a style to a paragraph, but styles can also be applied to ranges of text.</span></span>

    ```js
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

## <a name="apply-a-custom-style-to-text"></a><span data-ttu-id="34fc5-113">Aplicar um estilo personalizado ao texto</span><span class="sxs-lookup"><span data-stu-id="34fc5-113">Apply a custom style to text</span></span>

1. <span data-ttu-id="34fc5-114">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="34fc5-114">Open the file index.html.</span></span>
2. <span data-ttu-id="34fc5-115">Abaixo do `div` que contém o botão `apply-style`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="34fc5-115">Below the `div` that contains the `apply-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. <span data-ttu-id="34fc5-116">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="34fc5-116">Open the app.js file.</span></span>

4. <span data-ttu-id="34fc5-117">Abaixo da linha que atribui um identificador de clique ao botão `apply-style`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="34fc5-117">Below the line that assigns a click handler to the `apply-style` button, add the following code:</span></span>

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. <span data-ttu-id="34fc5-118">Abaixo da função `applyStyle`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="34fc5-118">Below the `applyStyle` function, add the following function:</span></span>

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

7. <span data-ttu-id="34fc5-119">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="34fc5-119">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="34fc5-120">O código aplica um estilo personalizado que ainda não existe.</span><span class="sxs-lookup"><span data-stu-id="34fc5-120">Note that the code applies a custom style that does not exist yet.</span></span> <span data-ttu-id="34fc5-121">Você criará um estilo com o nome **MyCustomStyle** na etapa [Testar o suplemento](#test-the-add-in).</span><span class="sxs-lookup"><span data-stu-id="34fc5-121">You'll create a style with the name **MyCustomStyle** in the [Test the add-in](#test-the-add-in) step.</span></span>

    ```js
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

## <a name="change-the-font-of-text"></a><span data-ttu-id="34fc5-122">Alterar a fonte do texto</span><span class="sxs-lookup"><span data-stu-id="34fc5-122">Change the font of text</span></span>

1. <span data-ttu-id="34fc5-123">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="34fc5-123">Open the file index.html.</span></span>
2. <span data-ttu-id="34fc5-124">Abaixo do `div` que contém o botão `apply-custom-style`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="34fc5-124">Below the `div` that contains the `apply-custom-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. <span data-ttu-id="34fc5-125">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="34fc5-125">Open the app.js file.</span></span>

4. <span data-ttu-id="34fc5-126">Abaixo da linha que atribui um identificador de clique ao botão `apply-custom-style`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="34fc5-126">Below the line that assigns a click handler to the `apply-custom-style` button, add the following code:</span></span>

    ```js
    $('#change-font').click(changeFont);
    ```

5. <span data-ttu-id="34fc5-127">Abaixo da função `applyCustomStyle`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="34fc5-127">Below the `applyCustomStyle` function, add the following function:</span></span>

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

7. <span data-ttu-id="34fc5-128">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="34fc5-128">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="34fc5-129">O código recebe uma referência para o segundo parágrafo usando o método `ParagraphCollection.getFirst` encadeado para o método `Paragraph.getNext`.</span><span class="sxs-lookup"><span data-stu-id="34fc5-129">Note that the code gets a reference to the second paragraph by using the `ParagraphCollection.getFirst` method chained to the `Paragraph.getNext` method.</span></span>

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="34fc5-130">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="34fc5-130">Test the add-in</span></span>

1. <span data-ttu-id="34fc5-131">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl+C duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="34fc5-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="34fc5-132">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="34fc5-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="34fc5-133">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="34fc5-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="34fc5-134">Para fazer isso, interrompa o processo do servidor para que o prompt apareça e você possa inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="34fc5-134">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="34fc5-135">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="34fc5-135">After the build, you restart the server.</span></span> <span data-ttu-id="34fc5-136">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="34fc5-136">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="34fc5-137">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.</span><span class="sxs-lookup"><span data-stu-id="34fc5-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="34fc5-138">Execute o comando `npm start` para iniciar um servidor Web em um localhost.</span><span class="sxs-lookup"><span data-stu-id="34fc5-138">Run the command `npm start` to start a web server running on localhost.</span></span>   
4. <span data-ttu-id="34fc5-139">Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="34fc5-139">Reload the task pane by closing it, and then on the **Home** menu select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="34fc5-140">Verifique se há pelo menos três parágrafos no documento.</span><span class="sxs-lookup"><span data-stu-id="34fc5-140">Be sure there are at least three paragraphs in the document.</span></span> <span data-ttu-id="34fc5-141">É possível escolher **Inserir Parágrafo** três vezes.</span><span class="sxs-lookup"><span data-stu-id="34fc5-141">You can choose **Insert Paragraph** three times.</span></span> <span data-ttu-id="34fc5-142">*Verifique com atenção se não há um parágrafo em branco no final do documento. Se houver, exclua-o.*</span><span class="sxs-lookup"><span data-stu-id="34fc5-142">*Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.*</span></span>
6. <span data-ttu-id="34fc5-143">No Word, crie um estilo personalizado chamado "MyCustomStyle".</span><span class="sxs-lookup"><span data-stu-id="34fc5-143">In Word, create a custom style named "MyCustomStyle".</span></span> <span data-ttu-id="34fc5-144">Pode ter a formatação que você quiser.</span><span class="sxs-lookup"><span data-stu-id="34fc5-144">It can have any formatting that you want.</span></span>
7. <span data-ttu-id="34fc5-145">Escolha o botão **Aplicar Estilo**.</span><span class="sxs-lookup"><span data-stu-id="34fc5-145">Choose the **Apply Style** button.</span></span> <span data-ttu-id="34fc5-146">O primeiro parágrafo receberá o estilo interno **Referência Intensa**.</span><span class="sxs-lookup"><span data-stu-id="34fc5-146">The first paragraph will be styled with the built-in style **Intense Reference**.</span></span>
8. <span data-ttu-id="34fc5-147">Escolha o botão **Aplicar Estilo Personalizado**.</span><span class="sxs-lookup"><span data-stu-id="34fc5-147">Choose the **Apply Custom Style** button.</span></span> <span data-ttu-id="34fc5-148">O último parágrafo receberá seu estilo personalizado.</span><span class="sxs-lookup"><span data-stu-id="34fc5-148">The last paragraph will be styled with your custom style.</span></span> <span data-ttu-id="34fc5-149">(Se parecer que nada acontece, talvez o último parágrafo esteja em branco.</span><span class="sxs-lookup"><span data-stu-id="34fc5-149">(If nothing seems to happen, the last paragraph might be blank.</span></span> <span data-ttu-id="34fc5-150">Se estiver, adicione um texto a ele).</span><span class="sxs-lookup"><span data-stu-id="34fc5-150">If so, add some text to it.)</span></span>
9. <span data-ttu-id="34fc5-151">Escolha o botão **Alterar Fonte**.</span><span class="sxs-lookup"><span data-stu-id="34fc5-151">Choose the **Change Font** button.</span></span> <span data-ttu-id="34fc5-152">A fonte do segundo parágrafo muda para 18 pt, negrito, Courier New.</span><span class="sxs-lookup"><span data-stu-id="34fc5-152">The font of the second paragraph changes to 18 pt., bold, Courier New.</span></span>

    ![Tutorial do Word: Aplicar estilos e fonte](../images/word-tutorial-apply-styles-and-font.png)
