<span data-ttu-id="38411-101">Nesta etapa do tutorial, você aprenderá a inserir imagens, HTML e tabelas no documento.</span><span class="sxs-lookup"><span data-stu-id="38411-101">In this step of the tutorial, you'll learn how to insert images, HTML, and tables into the document.</span></span>

> [!NOTE]
> <span data-ttu-id="38411-p101">Esta página descreve uma etapa individual de um tutorial de suplemento do Word. Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou por outro link direto, acesse a página de introdução do [tutorial de suplemento do Word](../tutorials/word-tutorial.yml) para começar pelo início.</span><span class="sxs-lookup"><span data-stu-id="38411-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="insert-an-image"></a><span data-ttu-id="38411-104">Inserir uma imagem</span><span class="sxs-lookup"><span data-stu-id="38411-104">Insert an image</span></span>

1. <span data-ttu-id="38411-105">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="38411-105">Open the project in your code editor.</span></span>
2. <span data-ttu-id="38411-106">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="38411-106">Open the file index.html.</span></span>
3. <span data-ttu-id="38411-107">Abaixo do `div` que contém o botão `replace-text`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="38411-107">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-image">Insert Image</button>
    </div>
    ```

4. <span data-ttu-id="38411-108">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="38411-108">Open the app.js file.</span></span>

5. <span data-ttu-id="38411-109">Na parte superior do arquivo, logo abaixo da linha use-strict, adicione a seguinte linha.</span><span class="sxs-lookup"><span data-stu-id="38411-109">Near the top of the file, just below the use-strict line, add the following line.</span></span> <span data-ttu-id="38411-110">Essa linha importa uma variável de outro arquivo.</span><span class="sxs-lookup"><span data-stu-id="38411-110">This line imports a variable from another file.</span></span> <span data-ttu-id="38411-111">A variável é uma cadeia de caracteres base 64 que codifica uma imagem.</span><span class="sxs-lookup"><span data-stu-id="38411-111">The variable is a base 64 string that encodes an image.</span></span> <span data-ttu-id="38411-112">Para ver a cadeia de caracteres codificada, abra o arquivo base64Image.js na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="38411-112">To see the encoded string, open the base64Image.js file in the root of the project.</span></span>

    ```js
    import { base64Image } from "./base64Image";
    ```

6. <span data-ttu-id="38411-113">Abaixo da linha que atribui um identificador de clique ao botão `replace-text`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="38411-113">Below the line that assigns a click handler to the `replace-text` button, add the following code:</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

7. <span data-ttu-id="38411-114">Abaixo da função `replaceText`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="38411-114">Below the `replaceText` function, add the following function:</span></span>

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

8. <span data-ttu-id="38411-115">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="38411-115">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="38411-116">Esta linha insere a imagem codificada em base 64 no final do documento.</span><span class="sxs-lookup"><span data-stu-id="38411-116">Note that this line inserts the base 64 encoded image at the end of the document.</span></span> <span data-ttu-id="38411-117">(O objeto `Paragraph` também tem um método `insertInlinePictureFromBase64` e outros métodos `insert*`.</span><span class="sxs-lookup"><span data-stu-id="38411-117">(The `Paragraph` object also has an `insertInlinePictureFromBase64` method and other `insert*` methods.</span></span> <span data-ttu-id="38411-118">Confira a seção insertHTML a seguir para conferir um exemplo).</span><span class="sxs-lookup"><span data-stu-id="38411-118">See the following insertHTML section for an example.)</span></span>

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

## <a name="insert-html"></a><span data-ttu-id="38411-119">Inserir HTML</span><span class="sxs-lookup"><span data-stu-id="38411-119">Insert HTML</span></span>

1. <span data-ttu-id="38411-120">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="38411-120">Open the file index.html.</span></span>
2. <span data-ttu-id="38411-121">Abaixo do `div` que contém o botão `insert-image`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="38411-121">Below the `div` that contains the `insert-image` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-html">Insert HTML</button>
    </div>
    ```

3. <span data-ttu-id="38411-122">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="38411-122">Open the app.js file.</span></span>

4. <span data-ttu-id="38411-123">Abaixo da linha que atribui um identificador de clique ao botão `insert-image`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="38411-123">Below the line that assigns a click handler to the `insert-image` button, add the following code:</span></span>

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. <span data-ttu-id="38411-124">Abaixo da função `insertImage`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="38411-124">Below the `insertImage` function, add the following function:</span></span>

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

6. <span data-ttu-id="38411-p104">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="38411-p104">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="38411-127">A primeira linha adiciona um parágrafo em branco ao final do documento.</span><span class="sxs-lookup"><span data-stu-id="38411-127">The first line adds a blank paragraph to the end of the document.</span></span> 
   - <span data-ttu-id="38411-128">A segunda linha insere uma cadeia de caracteres de HTML no final do parágrafo; especificamente dois parágrafos, um formatado com a fonte Verdana, e o outro com estilo padrão de documento do Word.</span><span class="sxs-lookup"><span data-stu-id="38411-128">The second line inserts a string of HTML at the end of the paragraph; specifically two paragraphs, one formatted with Verdana font, the other with the default styling of the Word document.</span></span> <span data-ttu-id="38411-129">(Conforme mostrado anteriormente no método `insertImage`, o objeto `context.document.body` também tem os métodos `insert*`).</span><span class="sxs-lookup"><span data-stu-id="38411-129">(As you saw in the `insertImage` method earlier, the `context.document.body` object also has the `insert*` methods.)</span></span>

    ```js
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

## <a name="insert-table"></a><span data-ttu-id="38411-130">Inserir Tabela</span><span class="sxs-lookup"><span data-stu-id="38411-130">Insert Table</span></span>

1. <span data-ttu-id="38411-131">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="38411-131">Open the file index.html.</span></span>
2. <span data-ttu-id="38411-132">Abaixo do `div` que contém o botão `insert-html`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="38411-132">Below the `div` that contains the `insert-html` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-table">Insert Table</button>
    </div>
    ```

3. <span data-ttu-id="38411-133">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="38411-133">Open the app.js file.</span></span>

4. <span data-ttu-id="38411-134">Abaixo da linha que atribui um identificador de clique ao botão `insert-html`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="38411-134">Below the line that assigns a click handler to the `insert-html` button, add the following code:</span></span>

    ```js
    $('#insert-table').click(insertTable);
    ```

5. <span data-ttu-id="38411-135">Abaixo da função `insertHTML`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="38411-135">Below the `insertHTML` function, add the following function:</span></span>

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

6. <span data-ttu-id="38411-136">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="38411-136">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="38411-137">Essa linha usa o método `ParagraphCollection.getFirst` para obter uma referência do primeiro parágrafo e, depois, usa o método `Paragraph.getNext` para obter uma referência para o segundo parágrafo.</span><span class="sxs-lookup"><span data-stu-id="38411-137">Note that this line uses the `ParagraphCollection.getFirst` method to get a reference ot the first paragraph and then uses the `Paragraph.getNext` method to get a reference to the second paragraph.</span></span>

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

7. <span data-ttu-id="38411-p107">Substitua `TODO2` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="38411-p107">Replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="38411-140">Os dois primeiros parâmetros do método `insertTable` especificam o número de linhas e colunas.</span><span class="sxs-lookup"><span data-stu-id="38411-140">The first two parameters of the `insertTable` method specify the number of rows and columns.</span></span>
   - <span data-ttu-id="38411-141">O terceiro parâmetro especifica onde inserir a tabela, nesse caso, depois do parágrafo.</span><span class="sxs-lookup"><span data-stu-id="38411-141">The third parameter specifies where to insert the table, in this case after the paragraph.</span></span>
   - <span data-ttu-id="38411-142">O quarto parâmetro é uma matriz bidimensional que define os valores das células da tabela.</span><span class="sxs-lookup"><span data-stu-id="38411-142">The fourth parameter is a two-dimensional array that sets the values of the table cells.</span></span>
   - <span data-ttu-id="38411-143">A tabela terá um estilo padrão simples, mas o método `insertTable` retornará um objeto `Table` com muitos membros, e alguns deles são usados para alterar o estilo de tabela.</span><span class="sxs-lookup"><span data-stu-id="38411-143">The table will have plain default styling, but the `insertTable` method returns a `Table` object with many members, some of which are used to style the table.</span></span>

    ```js
    const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="38411-144">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="38411-144">Test the add-in</span></span>


1. <span data-ttu-id="38411-145">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl+C duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="38411-145">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="38411-146">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="38411-146">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="38411-147">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="38411-147">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="38411-148">Para fazer isso, interrompa o processo do servidor para que o prompt apareça e você possa inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="38411-148">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="38411-149">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="38411-149">After the build, restart the server.</span></span> <span data-ttu-id="38411-150">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="38411-150">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="38411-151">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.</span><span class="sxs-lookup"><span data-stu-id="38411-151">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="38411-152">Execute o comando `npm start` para iniciar um servidor Web em um localhost.</span><span class="sxs-lookup"><span data-stu-id="38411-152">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="38411-153">Feche o painel de tarefas para recarregá-lo e, no menu **Página Inicial**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="38411-153">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="38411-154">No painel de tarefas, escolha **Inserir Parágrafo** pelo menos três vezes para garantir que haja alguns parágrafos no documento.</span><span class="sxs-lookup"><span data-stu-id="38411-154">In the taskpane, choose **Insert Paragraph** at least three times to ensure that there are a few paragraphs in the document.</span></span>
6. <span data-ttu-id="38411-155">Escolha o botão **Inserir Imagem**. Uma imagem é inserida no final do documento.</span><span class="sxs-lookup"><span data-stu-id="38411-155">Choose the **Insert Image** button and note that an image is inserted at the end of the document.</span></span>
7. <span data-ttu-id="38411-156">Escolha o botão **Inserir HTML**. Dois parágrafos são inseridos no final do documento, e o primeiro tem a fonte Verdana.</span><span class="sxs-lookup"><span data-stu-id="38411-156">Choose the **Insert HTML** button and note that two paragraphs are inserted at the end of the document, and that the first one has Verdana font.</span></span>
8. <span data-ttu-id="38411-157">Escolha o botão **Inserir Tabela**. Uma tabela é inserida após o segundo parágrafo.</span><span class="sxs-lookup"><span data-stu-id="38411-157">Choose the **Insert Table** button and note that a table is inserted after the second paragraph.</span></span>

    ![Tutorial do Word: Inserir imagem, HTML e tabela](../images/word-tutorial-insert-image-html-table.png)
