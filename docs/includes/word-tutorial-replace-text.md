<span data-ttu-id="a9524-101">Nesta etapa o tutorial, você adicionará texto dentro e fora dos intervalos de texto selecionados, e substituirá o texto de um intervalo selecionado.</span><span class="sxs-lookup"><span data-stu-id="a9524-101">In this step of the tutorial, you'll add text inside and outside of selected ranges of text, and replace the text of a selected range.</span></span> 

> [!NOTE]
> <span data-ttu-id="a9524-p101">Esta página descreve uma etapa individual de um tutorial de suplemento do Word. Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou por outro link direto, acesse a página de introdução do [tutorial de suplemento do Word](../tutorials/word-tutorial.yml) para começar pelo início.</span><span class="sxs-lookup"><span data-stu-id="a9524-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="add-text-inside-a-range"></a><span data-ttu-id="a9524-104">Adicionar texto dentro de um intervalo</span><span class="sxs-lookup"><span data-stu-id="a9524-104">Add text inside a range</span></span>

1. <span data-ttu-id="a9524-105">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="a9524-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="a9524-106">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="a9524-106">Open the file index.html.</span></span>
3. <span data-ttu-id="a9524-107">Abaixo do `div` que contém o botão `change-font`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="a9524-107">Below the `div` that contains the `change-font` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button>            
    </div>
    ```

4. <span data-ttu-id="a9524-108">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="a9524-108">Open the app.js file.</span></span>

5. <span data-ttu-id="a9524-109">Abaixo da linha que atribui um identificador de clique ao botão `change-font`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="a9524-109">Below the line that assigns a click handler to the `change-font` button, add the following code:</span></span>

    ```js
    $('#insert-text-into-range').click(insertTextIntoRange);
    ```

6. <span data-ttu-id="a9524-110">Abaixo da função `changeFont`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="a9524-110">Below the `changeFont` function, add the following function:</span></span>

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

7. <span data-ttu-id="a9524-p102">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="a9524-p102">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="a9524-113">o método serve para inserir a abreviação ["(C2R)"] no final do Intervalo cujo texto é "Clique para Executar".</span><span class="sxs-lookup"><span data-stu-id="a9524-113">The method is intended to insert the abbreviation ["(C2R)"] into the end of the Range whose text is "Click-to-Run".</span></span> <span data-ttu-id="a9524-114">Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.</span><span class="sxs-lookup"><span data-stu-id="a9524-114">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>
   - <span data-ttu-id="a9524-115">O primeiro parâmetro do método `Range.insertText` é a cadeia de caracteres a ser inserida no objeto `Range`.</span><span class="sxs-lookup"><span data-stu-id="a9524-115">The first parameter of the `Range.insertText` method is the string to insert into the `Range` object.</span></span>
   - <span data-ttu-id="a9524-116">O segundo parâmetro especifica onde no intervalo, o texto adicional deve ser inserido.</span><span class="sxs-lookup"><span data-stu-id="a9524-116">The second parameter specifies where in the range the additional text should be inserted.</span></span> <span data-ttu-id="a9524-117">Além de "Fim", as outras opções possíveis são "Início", "Antes", "Depois" e "Substituir".</span><span class="sxs-lookup"><span data-stu-id="a9524-117">Besides "End", the other possible options are "Start", "Before", "After", and "Replace".</span></span> 
   - <span data-ttu-id="a9524-118">A diferença entre "Fim" e "Depois" é que "Fim" insere o novo texto dentro o final do intervalo existente, mas "Depois" cria um novo intervalo com a cadeia de caracteres e insere o novo intervalo após o intervalo existente.</span><span class="sxs-lookup"><span data-stu-id="a9524-118">The difference between "End" and "After" is that "End" inserts the new text inside the end of the existing range, but "After" creates a new range with the string and inserts the new range after the existing range.</span></span> <span data-ttu-id="a9524-119">Da mesma forma, "Início" insere o texto dentro do início do intervalo existente, e "Antes" insere um novo intervalo.</span><span class="sxs-lookup"><span data-stu-id="a9524-119">Similarly, "Start" inserts text inside the beginning of the existing range and "Before" inserts a new range.</span></span> <span data-ttu-id="a9524-120">"Substituir" substitui o texto do intervalo existente pela cadeia de caracteres do primeiro parâmetro.</span><span class="sxs-lookup"><span data-stu-id="a9524-120">"Replace" replaces the text of the existing range with the string in the first parameter.</span></span>
   - <span data-ttu-id="a9524-121">Você viu em um estágio anterior do tutorial que os métodos insert\* do objeto de corpo não têm as opções "Antes" e "Depois".</span><span class="sxs-lookup"><span data-stu-id="a9524-121">You saw in an earlier stage of the tutorial that the insert\* methods of the body object do not have the "Before" and "After" options.</span></span> <span data-ttu-id="a9524-122">Isso ocorre porque não é possível colocar o conteúdo fora do corpo do documento.</span><span class="sxs-lookup"><span data-stu-id="a9524-122">This is because you can't put content outside of the document's body.</span></span>

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ``` 

8. <span data-ttu-id="a9524-123">Vamos deixar `TODO2` de lado até a próxima seção.</span><span class="sxs-lookup"><span data-stu-id="a9524-123">We'll skip over `TODO2` until the next section.</span></span> <span data-ttu-id="a9524-124">Substitua `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="a9524-124">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="a9524-125">Esse código é semelhante ao código que você criou no primeiro estágio do tutorial, exceto que, agora, você está inserindo um novo parágrafo no final do documento, em vez de no início.</span><span class="sxs-lookup"><span data-stu-id="a9524-125">This code is similar to the code you created in the first stage of the tutorial, except that now you are inserting a new paragraph at the end of the document instead of at the start.</span></span> <span data-ttu-id="a9524-126">Este novo parágrafo demonstrará que o novo texto agora faz parte do intervalo original.</span><span class="sxs-lookup"><span data-stu-id="a9524-126">This new paragraph will demonstrate that the new text is now part of the original range.</span></span>
 
    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text,
                             "End");
    ``` 

## <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="a9524-127">Adicione código para buscar propriedades do documento em objetos de script do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="a9524-127">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="a9524-128">Em todas as funções anteriores desta série de tutoriais, você colocou em fila comandos para *gravar* no documento do Office.</span><span class="sxs-lookup"><span data-stu-id="a9524-128">In all the previous functions in this series of tutorials, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="a9524-129">Cada função terminou com uma chamada para o método `context.sync()`, que envia os comandos em fila para o documento a ser executado.</span><span class="sxs-lookup"><span data-stu-id="a9524-129">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="a9524-130">Entretanto, o código adicionado na última etapa chama a propriedade `originalRange.text` e essa é uma grande diferença das funções anteriores que você escreveu, pois o objeto `originalRange` é apenas um objeto de proxy que existe no script do seu painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="a9524-130">But the code you added in the last step calls the `originalRange.text` property, and this is a significant difference from the earlier functions you wrote, because the `originalRange` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="a9524-131">Ele não sabe qual é o texto real do intervalo no documento, portanto, sua propriedade `text` não pode ter um valor real.</span><span class="sxs-lookup"><span data-stu-id="a9524-131">It doesn't know what the actual text of the range in the document is, so its `text` property can't have a real value.</span></span> <span data-ttu-id="a9524-132">Primeiro, é necessário buscar o valor de texto do intervalo no documento e usá-lo para definir o valor de `originalRange.text`.</span><span class="sxs-lookup"><span data-stu-id="a9524-132">It is necessary to first fetch the text value of the range from the document and use it to set the value of `originalRange.text`.</span></span> <span data-ttu-id="a9524-133">Somente então será possível chamar `originalRange.text` sem causar uma exceção.</span><span class="sxs-lookup"><span data-stu-id="a9524-133">Only then can `originalRange.text` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="a9524-134">Esse processo de busca tem três etapas:</span><span class="sxs-lookup"><span data-stu-id="a9524-134">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="a9524-135">Coloque em fila um comando para carregar (ou seja, fetch) as propriedades que seu código precisa ler.</span><span class="sxs-lookup"><span data-stu-id="a9524-135">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>
   2. <span data-ttu-id="a9524-136">Chame o método `sync` do objeto de contexto para enviar o comando em fila para o documento para execução e retornar as informações solicitadas.</span><span class="sxs-lookup"><span data-stu-id="a9524-136">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>
   3. <span data-ttu-id="a9524-137">Como o método `sync` é assíncrono, certifique-se de que ele tenha sido concluído antes que o código chame as propriedades que foram buscadas.</span><span class="sxs-lookup"><span data-stu-id="a9524-137">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="a9524-138">Estas etapas devem ser concluídas sempre que seu código precisar *ler* informações do documento do Office.</span><span class="sxs-lookup"><span data-stu-id="a9524-138">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="a9524-139">Substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="a9524-139">Replace `TODO2` with the following code.</span></span>
  
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

2. <span data-ttu-id="a9524-p109">Você não pode ter duas instruções `return` no mesmo caminho de código sem ramificações, portanto, exclua a linha final `return context.sync();` no final de `Word.run`. Você adicionará um novo final `context.sync` posteriormente neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="a9524-p109">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Word.run`. You'll add a new final `context.sync` later in this tutorial.</span></span> 
3. <span data-ttu-id="a9524-142">Recorte a linha `doc.body.insertParagraph` e cole no lugar de `TODO4`.</span><span class="sxs-lookup"><span data-stu-id="a9524-142">Cut the `doc.body.insertParagraph` line and paste in place of `TODO4`.</span></span> 
4. <span data-ttu-id="a9524-p110">Substitua `TODO5` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="a9524-p110">Replace `TODO5` with the following code. Note:</span></span>
   - <span data-ttu-id="a9524-145">Passar o método `sync` para uma função `then` garante que ele não seja executado até que a lógica `insertParagraph` tenha sido enfileirada.</span><span class="sxs-lookup"><span data-stu-id="a9524-145">Passing the `sync` method to a `then` function ensures that it does not run until the `insertParagraph` logic has been queued.</span></span>
   - <span data-ttu-id="a9524-146">O método `then` invoca qualquer função que é passada para ele e não é recomendável que `sync` seja chamado duas vezes, portanto, remova os "()" do fim de context.sync.</span><span class="sxs-lookup"><span data-stu-id="a9524-146">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of context.sync.</span></span>

    ```js
    .then(context.sync);
    ```

<span data-ttu-id="a9524-147">Quando terminar, a função inteira deve se parecer com o seguinte:</span><span class="sxs-lookup"><span data-stu-id="a9524-147">When you are done, the entire function should look like the following:</span></span>

  
```js
function insertTextIntoRange() {
    Word.run(function (context) {
        
        const doc = context.document;
        const originalRange = doc.getSelection();
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

## <a name="add-text-between-ranges"></a><span data-ttu-id="a9524-148">Adicionar texto entre intervalos</span><span class="sxs-lookup"><span data-stu-id="a9524-148">Add text between ranges</span></span>

1. <span data-ttu-id="a9524-149">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="a9524-149">Open the file index.html.</span></span>
2. <span data-ttu-id="a9524-150">Abaixo do `div` que contém o botão `insert-text-into-range`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="a9524-150">Below the `div` that contains the `insert-text-into-range` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button>            
    </div>
    ```

3. <span data-ttu-id="a9524-151">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="a9524-151">Open the app.js file.</span></span>

4. <span data-ttu-id="a9524-152">Abaixo da linha que atribui um identificador de clique ao botão `insert-text-into-range`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="a9524-152">Below the line that assigns a click handler to the `insert-text-into-range` button, add the following code:</span></span>

    ```js
    $('#insert-text-outside-range').click(insertTextBeforeRange);
    ```

5. <span data-ttu-id="a9524-153">Abaixo da função `insertTextIntoRange`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="a9524-153">Below the `insertTextIntoRange` function, add the following function:</span></span>

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

6. <span data-ttu-id="a9524-p111">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="a9524-p111">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="a9524-156">O método serve para adicionar um intervalo cujo texto seja "Office 2019", antes do intervalo com o texto "Office 365".</span><span class="sxs-lookup"><span data-stu-id="a9524-156">The method is intended to add a range whose text is "Office 2019, " before the range with text "Office 365".</span></span> <span data-ttu-id="a9524-157">Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.</span><span class="sxs-lookup"><span data-stu-id="a9524-157">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>
   - <span data-ttu-id="a9524-158">O primeiro parâmetro do método `Range.insertText` é a cadeia de caracteres a ser adicionada.</span><span class="sxs-lookup"><span data-stu-id="a9524-158">The first parameter of the `Range.insertText` method is the string to add.</span></span>
   - <span data-ttu-id="a9524-159">O segundo parâmetro especifica onde no intervalo, o texto adicional deve ser inserido.</span><span class="sxs-lookup"><span data-stu-id="a9524-159">The second parameter specifies where in the range the additional text should be inserted.</span></span> <span data-ttu-id="a9524-160">Para ter mais detalhes sobre as opções de local, confira a discussão anterior sobre a função `insertTextIntoRange`.</span><span class="sxs-lookup"><span data-stu-id="a9524-160">For more details about the location options, see the previous discussion of the `insertTextIntoRange` function.</span></span>

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ``` 

7. <span data-ttu-id="a9524-161">Substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="a9524-161">Replace `TODO2` with the following code.</span></span> 
 
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

8. <span data-ttu-id="a9524-162">Substitua `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="a9524-162">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="a9524-163">Este novo parágrafo demonstrará que o novo texto ***não*** faz parte do intervalo original selecionado.</span><span class="sxs-lookup"><span data-stu-id="a9524-163">This new paragraph will demonstrate the fact that the new text is ***not*** part of the original selected range.</span></span> <span data-ttu-id="a9524-164">O intervalo original ainda contém o texto que tinha quando foi selecionado.</span><span class="sxs-lookup"><span data-stu-id="a9524-164">The original range still has only the text it had when it was selected.</span></span>
 
    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                             "End");
    ``` 

9. <span data-ttu-id="a9524-165">Substitua `TODO4` pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="a9524-165">Replace `TODO4` with the following code:</span></span>

    ```js
    .then(context.sync);
    ```


## <a name="replace-the-text-of-a-range"></a><span data-ttu-id="a9524-166">Substitua o texto de um intervalo.</span><span class="sxs-lookup"><span data-stu-id="a9524-166">Replace the text of a range</span></span>

1. <span data-ttu-id="a9524-167">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="a9524-167">Open the file index.html.</span></span>
2. <span data-ttu-id="a9524-168">Abaixo do `div` que contém o botão `insert-text-outside-range`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="a9524-168">Below the `div` that contains the `insert-text-outside-range` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-text">Change Quantity Term</button>            
    </div>
    ```

3. <span data-ttu-id="a9524-169">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="a9524-169">Open the app.js file.</span></span>

4. <span data-ttu-id="a9524-170">Abaixo da linha que atribui um identificador de clique ao botão `insert-text-outside-range`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="a9524-170">Below the line that assigns a click handler to the `insert-text-outside-range` button, add the following code:</span></span>

    ```js
    $('#replace-text').click(replaceText);
    ```

5. <span data-ttu-id="a9524-171">Abaixo da função `insertTextBeforeRange`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="a9524-171">Below the `insertTextBeforeRange` function, add the following function:</span></span>

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

6. <span data-ttu-id="a9524-172">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="a9524-172">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="a9524-173">O método serve para substituir a cadeia de caracteres "várias" pela cadeia "muitos".</span><span class="sxs-lookup"><span data-stu-id="a9524-173">Note that the method is intended to replace the string "several" with the string "many".</span></span> <span data-ttu-id="a9524-174">Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.</span><span class="sxs-lookup"><span data-stu-id="a9524-174">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace"); 
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="a9524-175">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="a9524-175">Test the add-in</span></span>

1. <span data-ttu-id="a9524-176">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl + C duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="a9524-176">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="a9524-177">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="a9524-177">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="a9524-178">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="a9524-178">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="a9524-179">Para fazer isso, interrompa o processo do servidor para que o prompt apareça e você possa inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="a9524-179">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="a9524-180">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="a9524-180">After the build, restart the server.</span></span> <span data-ttu-id="a9524-181">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="a9524-181">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="a9524-182">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.</span><span class="sxs-lookup"><span data-stu-id="a9524-182">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="a9524-183">Execute o comando `npm start` para iniciar um servidor Web em um localhost.</span><span class="sxs-lookup"><span data-stu-id="a9524-183">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="a9524-184">Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="a9524-184">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="a9524-185">No painel de tarefas, escolha **Inserir Parágrafo** para garantir que haja um parágrafo no início do documento.</span><span class="sxs-lookup"><span data-stu-id="a9524-185">In the taskpane, choose **Insert Paragraph** to ensure that there is a paragraph at the start of the document.</span></span>
6. <span data-ttu-id="a9524-186">Selecione um texto.</span><span class="sxs-lookup"><span data-stu-id="a9524-186">Select some text.</span></span> <span data-ttu-id="a9524-187">Selecionar a frase "Clique para Executar" fará mais sentido.</span><span class="sxs-lookup"><span data-stu-id="a9524-187">Selecting the phrase "Click-to-Run" will make the most sense.</span></span> <span data-ttu-id="a9524-188">*Tenha cuidado para não incluir o espaço anterior ou seguinte na seleção.*</span><span class="sxs-lookup"><span data-stu-id="a9524-188">*Be careful not to include the preceding or following space in the selection.*</span></span>
7. <span data-ttu-id="a9524-189">Escolha o botão **Inserir Abreviação**.</span><span class="sxs-lookup"><span data-stu-id="a9524-189">Choose the **Insert Abbreviation** button.</span></span> <span data-ttu-id="a9524-190">"(C2R)" é adicionado.</span><span class="sxs-lookup"><span data-stu-id="a9524-190">Note that " (C2R)" is added.</span></span> <span data-ttu-id="a9524-191">Na parte inferior do documento, um novo parágrafo é adicionado com o texto inteiro expandido porque a nova cadeia de caracteres foi adicionada ao intervalo existente.</span><span class="sxs-lookup"><span data-stu-id="a9524-191">Note also that at the bottom of the document a new paragraph is added with the entire expanded text because the new string was added to the existing range.</span></span>
8. <span data-ttu-id="a9524-192">Selecione um texto.</span><span class="sxs-lookup"><span data-stu-id="a9524-192">Select some text.</span></span> <span data-ttu-id="a9524-193">Selecionar a frase "Office 365" fará mais sentido.</span><span class="sxs-lookup"><span data-stu-id="a9524-193">Selecting the phrase "Office 365" will make the most sense.</span></span> <span data-ttu-id="a9524-194">*Tenha cuidado para não incluir o espaço anterior ou seguinte na seleção.*</span><span class="sxs-lookup"><span data-stu-id="a9524-194">*Be careful not to include the preceding or following space in the selection.*</span></span>
9. <span data-ttu-id="a9524-195">Escolha o botão **Adicionar Informações de Versão**.</span><span class="sxs-lookup"><span data-stu-id="a9524-195">Choose the **Add Version Info** button.</span></span> <span data-ttu-id="a9524-196">"Office 2019" está inserido entre "Office 2016" e "Office 365".</span><span class="sxs-lookup"><span data-stu-id="a9524-196">Note that "Office 2019, " is inserted between "Office 2016" and "Office 365".</span></span> <span data-ttu-id="a9524-197">Na parte inferior do documento um novo parágrafo foi adicionado, mas ele contém apenas o texto selecionado originalmente porque a nova cadeia de caracteres tornou-se um intervalo novo, em vez de ser adicionada ao intervalo original.</span><span class="sxs-lookup"><span data-stu-id="a9524-197">Note also that at the bottom of the document a new paragraph is added but it contains only the originally selected text because the new string became a new range rather than being added to the original range.</span></span>
10. <span data-ttu-id="a9524-198">Selecione um texto.</span><span class="sxs-lookup"><span data-stu-id="a9524-198">Select some text.</span></span> <span data-ttu-id="a9524-199">Selecionar a palavra "vários" fará mais sentido.</span><span class="sxs-lookup"><span data-stu-id="a9524-199">Selecting the word "several" will make the most sense.</span></span> <span data-ttu-id="a9524-200">*Tenha cuidado para não incluir o espaço anterior ou seguinte na seleção.*</span><span class="sxs-lookup"><span data-stu-id="a9524-200">*Be careful not to include the preceding or following space in the selection.*</span></span>
11. <span data-ttu-id="a9524-201">Escolha o botão **Alterar Termo de Quantidade**.</span><span class="sxs-lookup"><span data-stu-id="a9524-201">Choose the **Change Quantity Term** button.</span></span> <span data-ttu-id="a9524-202">"muitos" substitui o texto selecionado.</span><span class="sxs-lookup"><span data-stu-id="a9524-202">Note that "many" replaces the selected text.</span></span>

    ![Tutorial do Word: texto adicionado e substituído](../images/word-tutorial-text-replace.png)
