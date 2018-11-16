<span data-ttu-id="1e513-101">Nesta etapa do tutorial, você adicionará outro botão à faixa de opções que, quando selecionado, executa uma função que você precisará definir para ativar e desativar a proteção da planilha.</span><span class="sxs-lookup"><span data-stu-id="1e513-101">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

> [!NOTE]
> <span data-ttu-id="1e513-102">Esta página descreve uma etapa individual do tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="1e513-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="1e513-103">Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou via outro link direto, acesse a página de Introdução do [tutorial de suplemento do Excel](../tutorials/excel-tutorial.yml) para começá-lo do início.</span><span class="sxs-lookup"><span data-stu-id="1e513-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="1e513-104">Configure o manifesto para adicionar um segundo botão à faixa de opções</span><span class="sxs-lookup"><span data-stu-id="1e513-104">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="1e513-105">Abra o arquivo de manifesto **my-office-add-in-manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="1e513-105">Open the manifest file **my-office-add-in-manifest.xml**.</span></span>
2. <span data-ttu-id="1e513-106">Encontre o elemento `<Control>`.</span><span class="sxs-lookup"><span data-stu-id="1e513-106">Find the `<Control>` element.</span></span> <span data-ttu-id="1e513-107">Esse elemento define o botão **Mostrar Painel de Tarefas** na faixa de opções **Início** que você usa para iniciar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="1e513-107">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="1e513-108">Vamos adicionar um segundo botão ao mesmo grupo na faixa de opções **Início**.</span><span class="sxs-lookup"><span data-stu-id="1e513-108">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="1e513-109">Entre a marca de Controle final (`</Control>`) e a marca de Grupo final (`</Group>`), adicione a marcação a seguir.</span><span class="sxs-lookup"><span data-stu-id="1e513-109">In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. <span data-ttu-id="1e513-110">Substitua `TODO1` por uma cadeia de caracteres que fornece ao botão uma ID exclusiva no arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="1e513-110">Replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="1e513-111">Há apenas um outro botão no manifesto, portanto, isso não é difícil.</span><span class="sxs-lookup"><span data-stu-id="1e513-111">There's only one other button in the manifest, so this isn't difficult.</span></span> <span data-ttu-id="1e513-112">Como nosso botão ativará ou desativará a proteção da planilha, use "ToggleProtection".</span><span class="sxs-lookup"><span data-stu-id="1e513-112">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="1e513-113">Quando terminar, a marca de Controle de início inteira deve se parecer com o seguinte:</span><span class="sxs-lookup"><span data-stu-id="1e513-113">When you are done, the entire start Control tag should look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="1e513-114">Os próximos três `TODO`s definem “resid”, que significa ID de recurso.</span><span class="sxs-lookup"><span data-stu-id="1e513-114">The next three `TODO`s set "resid"s, which is short for resource ID.</span></span> <span data-ttu-id="1e513-115">Um recurso é uma cadeia de caracteres e você criará essas três cadeias de caracteres em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="1e513-115">A resource is a string, and you'll create these three strings in a later step.</span></span> <span data-ttu-id="1e513-116">Por enquanto, você precisa fornecer IDs aos recursos.</span><span class="sxs-lookup"><span data-stu-id="1e513-116">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="1e513-117">O rótulo do botão deve ser "Toggle Protection", mas a *ID* dessa cadeia de caracteres será "ProtectionButtonLabel", de forma que o elemento `Label` completo deve se parecer com o código a seguir:</span><span class="sxs-lookup"><span data-stu-id="1e513-117">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the completed `Label` element should look like the following code:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="1e513-118">O elemento `SuperTip` define a dica de ferramenta do botão.</span><span class="sxs-lookup"><span data-stu-id="1e513-118">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="1e513-119">O título da dica de ferramenta deve ser o mesmo que o rótulo do botão, por isso, usamos a mesma ID de recurso: "ProtectionButtonLabel".</span><span class="sxs-lookup"><span data-stu-id="1e513-119">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="1e513-120">A descrição da dica de ferramenta será "Click to turn protection of the worksheet on and off".</span><span class="sxs-lookup"><span data-stu-id="1e513-120">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="1e513-121">Mas o `ID` será "ProtectionButtonToolTip".</span><span class="sxs-lookup"><span data-stu-id="1e513-121">But the `ID` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="1e513-122">Portanto, quando terminar, a marcação `SuperTip` inteira deve se parecer com o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="1e513-122">So, when you are done, the whole `SuperTip` markup should look like the following code:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="1e513-123">Em um suplemento de produção,não é recomendável usar o mesmo ícone para dois botões diferentes; mas, para simplificar este tutorial, faremos isso.</span><span class="sxs-lookup"><span data-stu-id="1e513-123">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that.</span></span> <span data-ttu-id="1e513-124">Portanto, a marcação `Icon` em nosso novo `Control` é apenas uma cópia do elemento `Icon` do `Control` existente.</span><span class="sxs-lookup"><span data-stu-id="1e513-124">So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="1e513-125">O elemento `Action` dentro do elemento `Control` original já está presente no manifesto, tem seu tipo definido como `ShowTaskpane`, mas nosso novo botão não abrirá um painel de tarefas, mas sim executará uma função personalizada criada em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="1e513-125">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="1e513-126">Portanto, substitua `TODO5` por `ExecuteFunction`, que é o tipo de ação para botões que acionam funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="1e513-126">So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="1e513-127">A marca `Action` de início deve ser similar ao código abaixo:</span><span class="sxs-lookup"><span data-stu-id="1e513-127">The start `Action` tag should look like the following code:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="1e513-128">O elemento `Action` original tem elementos filhos que especificam uma ID do painel de tarefas e uma URL da página que deve ser aberta no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="1e513-128">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane.</span></span> <span data-ttu-id="1e513-129">No entanto, um elemento `Action` do tipo `ExecuteFunction` tem um único elemento filho que nomeia a função executada pelo controle.</span><span class="sxs-lookup"><span data-stu-id="1e513-129">But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes.</span></span> <span data-ttu-id="1e513-130">Você criará essa função em uma etapa posterior e ela será chamada de `toggleProtection`.</span><span class="sxs-lookup"><span data-stu-id="1e513-130">You'll create that function in a later step, and it will be called `toggleProtection`.</span></span> <span data-ttu-id="1e513-131">Então, substitua `TODO6` pela marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="1e513-131">So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="1e513-132">A marcação `Control` inteira deve ter a aparência a seguir:</span><span class="sxs-lookup"><span data-stu-id="1e513-132">The entire `Control` markup should now look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. <span data-ttu-id="1e513-133">Role para baixo até a seção `Resources` do manifesto.</span><span class="sxs-lookup"><span data-stu-id="1e513-133">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="1e513-134">Adicione a seguinte marcação como filho do elemento `bt:ShortStrings`.</span><span class="sxs-lookup"><span data-stu-id="1e513-134">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="1e513-135">Adicione a seguinte marcação como filho do elemento `bt:LongStrings`.</span><span class="sxs-lookup"><span data-stu-id="1e513-135">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="1e513-136">Não deixe de salvar o arquivo.</span><span class="sxs-lookup"><span data-stu-id="1e513-136">Be sure to save the file.</span></span>

## <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="1e513-137">Criar a função que protege a planilha</span><span class="sxs-lookup"><span data-stu-id="1e513-137">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="1e513-138">Abra o arquivo \function-file\function-file.js.</span><span class="sxs-lookup"><span data-stu-id="1e513-138">Open the file \function-file\function-file.js.</span></span>

2. <span data-ttu-id="1e513-139">O arquivo já tem uma Expressão de Função Invocada Imediatamente (IFFE).</span><span class="sxs-lookup"><span data-stu-id="1e513-139">The file already has an Immediately Invoked Function Expression (IFFE).</span></span> <span data-ttu-id="1e513-140">Não é necessário ter uma lógica de inicialização personalizada, portanto, deixe a função atribuída a `Office.initialize` com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="1e513-140">No custom initialization logic is needed, so leave the function that is assigned to `Office.initialize` with an empty body.</span></span> <span data-ttu-id="1e513-141">(Mas não a exclua.</span><span class="sxs-lookup"><span data-stu-id="1e513-141">(But do not delete it.</span></span> <span data-ttu-id="1e513-142">A propriedade `Office.initialize` não pode ser nula ou indefinida.) *Fora da IIFE*, adicione o seguinte código.</span><span class="sxs-lookup"><span data-stu-id="1e513-142">The `Office.initialize` property cannot be null or undefined.) *Outside of the IIFE*, add the following code.</span></span> <span data-ttu-id="1e513-143">Observe que é possível especificar um parâmetro `args` para o método e a última linha do método chamará `args.completed`.</span><span class="sxs-lookup"><span data-stu-id="1e513-143">Note that we specify an `args` parameter to the method and the very last line of the method calls `args.completed`.</span></span> <span data-ttu-id="1e513-144">Esse é um requisito para todos os comandos de suplemento do tipo **ExecuteFunction**.</span><span class="sxs-lookup"><span data-stu-id="1e513-144">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="1e513-145">Ele sinaliza para o aplicativo host do Office que a função terminou e que a interface do usuário podem ficar responsiva novamente.</span><span class="sxs-lookup"><span data-stu-id="1e513-145">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

3. <span data-ttu-id="1e513-146">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="1e513-146">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="1e513-147">O código usa propriedade de proteção do objeto de planilha em um padrão de botão de alternância padrão.</span><span class="sxs-lookup"><span data-stu-id="1e513-147">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="1e513-148">O `TODO2` será explicado na próxima seção.</span><span class="sxs-lookup"><span data-stu-id="1e513-148">The `TODO2` will be explained in the next section.</span></span>

    ```javascript
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

     if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ``` 

## <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="1e513-149">Adicione código para buscar propriedades do documento em objetos de script do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="1e513-149">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="1e513-150">Em todas as funções anteriores desta série de tutoriais, você colocou em fila comandos para *gravar* no documento do Office.</span><span class="sxs-lookup"><span data-stu-id="1e513-150">In all the earlier functions in this series of tutorials, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="1e513-151">Cada função terminou com uma chamada para o método `context.sync()`, que envia os comandos em fila para o documento a ser executado.</span><span class="sxs-lookup"><span data-stu-id="1e513-151">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="1e513-152">Entretanto, o código adicionado na última etapa chama a propriedade `sheet.protection.protected` e essa é uma grande diferença das funções anteriores que você escreveu, pois o objeto `sheet` é apenas um objeto de proxy que existe no script do seu painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="1e513-152">But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="1e513-153">Ele não sabe qual é o estado real de proteção do documento, portanto, sua propriedade `protection.protected` não pode ter um valor real.</span><span class="sxs-lookup"><span data-stu-id="1e513-153">It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="1e513-154">É necessário primeiro buscar o status de proteção do documento e definir o valor de `sheet.protection.protected`.</span><span class="sxs-lookup"><span data-stu-id="1e513-154">It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="1e513-155">Somente então será possível chamar `sheet.protection.protected` sem causar uma exceção.</span><span class="sxs-lookup"><span data-stu-id="1e513-155">Only then can `sheet.protection.protected` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="1e513-156">Esse processo de busca tem três etapas:</span><span class="sxs-lookup"><span data-stu-id="1e513-156">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="1e513-157">Coloque em fila um comando para carregar (ou seja, fetch) as propriedades que seu código precisa ler.</span><span class="sxs-lookup"><span data-stu-id="1e513-157">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>
   2. <span data-ttu-id="1e513-158">Chame o método `sync` do objeto de contexto para enviar o comando em fila para o documento para execução e retornar as informações solicitadas.</span><span class="sxs-lookup"><span data-stu-id="1e513-158">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>
   3. <span data-ttu-id="1e513-159">Como o método `sync` é assíncrono, certifique-se de que ele tenha sido concluído antes que o código chame as propriedades que foram buscadas.</span><span class="sxs-lookup"><span data-stu-id="1e513-159">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="1e513-160">Essas etapas devem ser concluídas sempre que seu código precisar *ler* informações do documento do Office.</span><span class="sxs-lookup"><span data-stu-id="1e513-160">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="1e513-p112">Na função `toggleProtection`, substitua `TODO2` pelo seguinte código. Observação:</span><span class="sxs-lookup"><span data-stu-id="1e513-p112">In the `toggleProtection` function, replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="1e513-163">Todos os objetos do Excel têm um método `load`.</span><span class="sxs-lookup"><span data-stu-id="1e513-163">Every Excel object has a `load` method.</span></span> <span data-ttu-id="1e513-164">Especifique as propriedades do objeto que você deseja ler no parâmetro como uma cadeia de caracteres de nomes delimitados por vírgulas.</span><span class="sxs-lookup"><span data-stu-id="1e513-164">You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names.</span></span> <span data-ttu-id="1e513-165">Nesse caso, a propriedade que você precisa ler é uma subpropriedade de `protection`.</span><span class="sxs-lookup"><span data-stu-id="1e513-165">In this case, the property you need to read is a subproperty of the `protection` property.</span></span> <span data-ttu-id="1e513-166">Referencie a subpropriedade quase exatamente como você faria em qualquer lugar do seu código, mas usando uma barra (“/”) em vez de um ponto (".").</span><span class="sxs-lookup"><span data-stu-id="1e513-166">You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>
   - <span data-ttu-id="1e513-167">Para garantir que a lógica de botão de alternância, `sheet.protection.protected`, não seja executada até após `sync` ser concluído e o `sheet.protection.protected` ser atribuída ao valor correto buscado no documento, ele será movido (na próxima etapa) para uma função `then` que não será executada até `sync` ser concluído.</span><span class="sxs-lookup"><span data-stu-id="1e513-167">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

    ```javascript
    sheet.load('protection/protected');
    return context.sync()
        .then(
            function() {
                // TODO3: Move the queued toggle logic here.
            }
        )
        // TODO4: Move the final call of `context.sync` here and ensure that it
        //        does not run until the toggle logic has been queued.
    ``` 

2. <span data-ttu-id="1e513-168">Você não pode ter duas instruções `return` no mesmo caminho de código sem ramificações, portanto, exclua a linha final `return context.sync();` no final de `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="1e513-168">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`.</span></span> <span data-ttu-id="1e513-169">Você adicionará um novo `context.sync` final em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="1e513-169">You will add a new final `context.sync`, in a later step.</span></span>
3. <span data-ttu-id="1e513-170">Recorte a estrutura `if ... else` na função `toggleProtection` e a cole no lugar de `TODO3`.</span><span class="sxs-lookup"><span data-stu-id="1e513-170">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>
4. <span data-ttu-id="1e513-p115">Substitua `TODO4` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="1e513-p115">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="1e513-173">Passar o método `sync` para uma função `then` garante que ele não seja executado até que `sheet.protection.unprotect()` ou `sheet.protection.protect()` seja enfileirado.</span><span class="sxs-lookup"><span data-stu-id="1e513-173">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>
   - <span data-ttu-id="1e513-174">O método `then` invoca qualquer função que é passada para ele e não é recomendável que `sync` seja chamado duas vezes, portanto, remova os “()” do fim de `context.sync`.</span><span class="sxs-lookup"><span data-stu-id="1e513-174">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```javascript
    .then(context.sync);
    ```

   <span data-ttu-id="1e513-175">Quando terminar, a função inteira deve se parecer com o seguinte:</span><span class="sxs-lookup"><span data-stu-id="1e513-175">When you are done, the entire function should look like the following:</span></span>

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {            
          const sheet = context.workbook.worksheets.getActiveWorksheet();          
          sheet.load('protection/protected');

          return context.sync()
              .then(
                  function() {
                    if (sheet.protection.protected) {
                        sheet.protection.unprotect();
                    } else {
                        sheet.protection.protect();
                    }
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
        args.completed();
    }
    ```


## <a name="configure-the-script-loading-html-file"></a><span data-ttu-id="1e513-176">Configure o arquivo HTML de carregamento de script</span><span class="sxs-lookup"><span data-stu-id="1e513-176">Configure the script-loading HTML file</span></span>

<span data-ttu-id="1e513-177">Abra o arquivo /function-file/function-file.html.</span><span class="sxs-lookup"><span data-stu-id="1e513-177">Open the /function-file/function-file.html file.</span></span> <span data-ttu-id="1e513-178">Esse é um arquivo HTML sem IU que é chamado quando o usuário pressiona o botão **Ativar/Desativar Proteção da Planilha**.</span><span class="sxs-lookup"><span data-stu-id="1e513-178">This is a UI-less HTML file that is called when the user presses the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="1e513-179">O objetivo é carregar o método JavaScript que deve ser executado quando botão é pressionado.</span><span class="sxs-lookup"><span data-stu-id="1e513-179">Its purpose is to load the JavaScript method that should run when the button is pushed.</span></span> <span data-ttu-id="1e513-180">Esse arquivo não será alterado.</span><span class="sxs-lookup"><span data-stu-id="1e513-180">You are not going to change this file.</span></span> <span data-ttu-id="1e513-181">Basta observar que a segunda marca `<script>` carrega o functionfile.js.</span><span class="sxs-lookup"><span data-stu-id="1e513-181">Simply note that the second `<script>` tag loads the functionfile.js.</span></span>

   > [!NOTE]
   > <span data-ttu-id="1e513-182">O arquivo function-file.html e o arquivo function-file.js carregado são executados em um processo do IE completamente separado de painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="1e513-182">The function-file.html file and the function-file.js file that it loads run in an entirely separate IE process from the add-in's task pane.</span></span> <span data-ttu-id="1e513-183">Se o function-file.js foi transcompilado no mesmo arquivo bundle.js que o arquivo app.js, o suplemento precisará carregar duas cópias do arquivo bundle.js, o que anule o propósito do agrupamento.</span><span class="sxs-lookup"><span data-stu-id="1e513-183">If the function-file.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="1e513-184">Além disso, o arquivo function-file.js não contém qualquer JavaScript incompatível com o Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="1e513-184">In addition, the function-file.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="1e513-185">Por esses dois motivos, esse suplemento não transcompila o function-file.js.</span><span class="sxs-lookup"><span data-stu-id="1e513-185">For these two reasons, this add-in does not transpile the function-file.js at all.</span></span> 

## <a name="test-the-add-in"></a><span data-ttu-id="1e513-186">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="1e513-186">Test the add-in</span></span>

1. <span data-ttu-id="1e513-187">Feche todos os aplicativos do Office, incluindo o Excel.</span><span class="sxs-lookup"><span data-stu-id="1e513-187">Close all Office applications, including Excel.</span></span> 
2. <span data-ttu-id="1e513-188">Para excluir o cache do Office, exclua o conteúdo da pasta de cache.</span><span class="sxs-lookup"><span data-stu-id="1e513-188">Delete the Office cache by deleting the contents of the cache folder.</span></span> <span data-ttu-id="1e513-189">Isso é necessário para limpar totalmente a versão anterior do suplemento do host.</span><span class="sxs-lookup"><span data-stu-id="1e513-189">This is necessary to completely clear the old version of the add-in from the host.</span></span> 
    - <span data-ttu-id="1e513-190">No Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="1e513-190">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>
    - <span data-ttu-id="1e513-191">No Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="1e513-191">For Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>
3. <span data-ttu-id="1e513-192">Se, por algum motivo, o servidor não estiver executando, em uma janela do Git Bash ou em um prompt do sistema habilitado para Node.JS, acesse a pasta **Iniciar** do projeto e execute o comando `npm start`.</span><span class="sxs-lookup"><span data-stu-id="1e513-192">If for any reason, your server is not running, then in a Git Bash window, or Node.JS-enabled system prompt, navigate to the **Start** folder of the project and run the command `npm start`.</span></span> <span data-ttu-id="1e513-193">Não é necessário recriar o projeto, pois o único arquivo JavaScript que você alterou não faz parte do bundle.js interno.</span><span class="sxs-lookup"><span data-stu-id="1e513-193">You do not need to rebuild the project because the only JavaScript file you changed is not part of the built bundle.js.</span></span>
4. <span data-ttu-id="1e513-194">Usando a nova versão do arquivo de manifesto alterado, repita o processo de sideloading usando um dos seguintes métodos.</span><span class="sxs-lookup"><span data-stu-id="1e513-194">Using the new version of the changed manifest file, repeat the sideloading process by using one of the following methods.</span></span> <span data-ttu-id="1e513-195">*Você deve substituir a cópia anterior do arquivo de manifesto.*</span><span class="sxs-lookup"><span data-stu-id="1e513-195">*You should overwrite the previous copy of the manifest file.*</span></span>
    - <span data-ttu-id="1e513-196">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="1e513-196">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="1e513-197">Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="1e513-197">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="1e513-198">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="1e513-198">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
7. <span data-ttu-id="1e513-199">Abra qualquer planilha no Excel.</span><span class="sxs-lookup"><span data-stu-id="1e513-199">Open any worksheet in Excel.</span></span>
8. <span data-ttu-id="1e513-p121">Na Faixa de Opções, em **Página Inicial**, escolha **Ativar Proteger Planilha**. Observe que a maioria dos controles na Faixa de Opções está desabilitada (e visualmente esmaecida) conforme mostrado na captura de tela abaixo.</span><span class="sxs-lookup"><span data-stu-id="1e513-p121">On the **Home** ribbon, choose **Toggle Worksheet Protection**. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in screenshot below.</span></span> 
9. <span data-ttu-id="1e513-202">Escolha uma célula como se quisesse alterar o conteúdo.</span><span class="sxs-lookup"><span data-stu-id="1e513-202">Choose a cell as you would if you wanted to change its content.</span></span> <span data-ttu-id="1e513-203">Você receberá um erro informando que a planilha está protegida.</span><span class="sxs-lookup"><span data-stu-id="1e513-203">You get an error telling you that the worksheet is protected.</span></span>
10. <span data-ttu-id="1e513-204">Escolha **Ativar/Desativar Proteção da Planilha** novamente e os controles serão reabilitados e você poderá alterar os valores das células.</span><span class="sxs-lookup"><span data-stu-id="1e513-204">Choose **Toggle Worksheet Protection** again, and the controls are reenabled, and you can change cell values again.</span></span>

    ![Tutorial do Excel - Faixa de Opções com a Proteção Ativada](../images/excel-tutorial-ribbon-with-protection-on.png)
