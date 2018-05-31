<span data-ttu-id="595e9-101">Nesta etapa final do tutorial, você abre uma caixa de diálogo no suplemento, passa uma mensagem do processo de caixa de diálogo para o processo de painel de tarefas e fecha a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="595e9-101">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog.</span></span> <span data-ttu-id="595e9-102">As caixas de diálogo do Suplemento do Office são *não modais*: o usuário pode continuar a interagir com o documento no aplicativo do Office do host e com a página host no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="595e9-102">Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="595e9-103">Esta página descreve uma etapa individual do tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="595e9-103">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="595e9-104">Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou via outro link direto, acesse a página de Introdução do [tutorial de suplemento do Excel](../tutorials/excel-tutorial.yml) para começá-lo do início.</span><span class="sxs-lookup"><span data-stu-id="595e9-104">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="create-the-dialog-page"></a><span data-ttu-id="595e9-105">Crie a página da caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="595e9-105">Create the dialog page</span></span>

1. <span data-ttu-id="595e9-106">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="595e9-106">Open the project in your code editor.</span></span>
2. <span data-ttu-id="595e9-107">Crie um arquivo chamado popup.html na raiz do projeto (onde se encontra index.html).</span><span class="sxs-lookup"><span data-stu-id="595e9-107">Create a file in the root of the project (where index.html is) called popup.html.</span></span>
3. <span data-ttu-id="595e9-p103">Adicione a marcação a seguir em popup.html. Observação:</span><span class="sxs-lookup"><span data-stu-id="595e9-p103">Add the following markup to popup.html. Note:</span></span>
   - <span data-ttu-id="595e9-110">A página tem um `<input>` onde o usuário insere seu nome e um botão que enviará o nome para a página no painel de tarefas onde ele será exibido.</span><span class="sxs-lookup"><span data-stu-id="595e9-110">The page has a `<input>` where the user will enter his or her name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>
   - <span data-ttu-id="595e9-111">A marcação carrega um script chamado popup.js que você criará em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="595e9-111">The markup loads a script called popup.js that you will create in a later step.</span></span>
   - <span data-ttu-id="595e9-112">Ela também carrega uma biblioteca Office.JS e jQuery porque elas serão usadas em popup.js.</span><span class="sxs-lookup"><span data-stu-id="595e9-112">It also loads the Office.JS library and jQuery because they will be used in popup.js.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
        
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.min.css" />
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.components.css" />
            <link rel="stylesheet" href="app.css">
    
            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
            <script type="text/javascript" src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.1.min.js"></script>
            <script type="text/javascript" src="popup.js"></script>
    
        </head>
         <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
         <div class="padding">
            <p class="ms-font-xl">ENTER YOUR NAME</p>
         </div>        
        <div class="padding">
            <input id="name-box" type="text"/>
        <div>
        <div class="padding">
            <button id="ok-button" class="ms-Button">OK</button>
        </div>
    </body>
    </html>
    ```

4. <span data-ttu-id="595e9-113">Crie um arquivo chamado popup.js na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="595e9-113">Create a file in the root of the project called popup.js.</span></span>
5. <span data-ttu-id="595e9-p104">Adicione o código a seguir ao popup.js. Observação:</span><span class="sxs-lookup"><span data-stu-id="595e9-p104">Add the following code to popup.js. Note:</span></span>
   - <span data-ttu-id="595e9-116">*Todas as páginas que chamam APIs na biblioteca Office.JS devem atribuir uma função à propriedade `Office.initialize`.*</span><span class="sxs-lookup"><span data-stu-id="595e9-116">*Every page that calls APIs in the Office.JS library must assign a function to the `Office.initialize` property.*</span></span> <span data-ttu-id="595e9-117">Se nenhuma inicialização for necessária, a função poderá ter um corpo vazio, mas a propriedade não deve ser deixada indefinida, atribuída a nulo ou a um valor que não seja uma função.</span><span class="sxs-lookup"><span data-stu-id="595e9-117">If no initialization is needed, then the function can have an empty body, but the property must not be left undefined, assigned to null or to a non-function value.</span></span> <span data-ttu-id="595e9-118">Por exemplo, veja o arquivo app.js na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="595e9-118">For an example, see the app.js file in the project root.</span></span> <span data-ttu-id="595e9-119">O código que cria a tarefa deve ser executado antes de qualquer chamada para Office.JS; por isso, a tarefa se encontra em um arquivo de script que é carregado pela página, como neste caso.</span><span class="sxs-lookup"><span data-stu-id="595e9-119">The code that makes the assignment must run before any calls to Office.JS; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>
   - <span data-ttu-id="595e9-p106">A função jQuery `ready` é chamada dentro do método `initialize`. É uma regra quase universal que o código de carregamento, inicialização ou bootstrapping de outras bibliotecas JavaScript deva estar dentro da função `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="595e9-p106">The jQuery `ready` function is called inside the `initialize` method. It is an almost universal rule that the loading, initializing, or bootstrapping code of other JavaScript libraries should be inside the `Office.initialize` function.</span></span>

    ```js
    (function () {
    "use strict";

        Office.initialize = function() {        
            $(document).ready(function () {  
    
                // TODO1: Assign handler to the OK button.
    
            });
        }

        // TODO2: Create the OK button handler
    
    }());    
    ```

6. <span data-ttu-id="595e9-122">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="595e9-122">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="595e9-123">Você criará a função `sendStringToParentPage` na próxima etapa.</span><span class="sxs-lookup"><span data-stu-id="595e9-123">You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. <span data-ttu-id="595e9-124">Substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="595e9-124">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="595e9-125">O método `messageParent` passa seu parâmetro para a página pai, neste caso, a página no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="595e9-125">The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane.</span></span> <span data-ttu-id="595e9-126">O parâmetro pode ser um booliano ou uma cadeia de caracteres, que inclui tudo o que pode ser serializado como uma cadeia de caracteres, como XML ou JSON.</span><span class="sxs-lookup"><span data-stu-id="595e9-126">The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span> 

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. <span data-ttu-id="595e9-127">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="595e9-127">Save the file.</span></span>

   > [!NOTE]
   > <span data-ttu-id="595e9-128">O arquivo popup.html e o arquivo popup.js carregado são executados em um processo do Internet Explorer completamente separado de painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="595e9-128">The popup.html file, and the popup.js file that it loads, run in an entirely separate Internet Explorer process from the add-in's task pane.</span></span> <span data-ttu-id="595e9-129">Se o popup.js foi transcompilado no mesmo arquivo bundle.js que o arquivo app.js, o suplemento precisará carregar duas cópias do arquivo bundle.js, o que anule o propósito do agrupamento.</span><span class="sxs-lookup"><span data-stu-id="595e9-129">If the popup.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="595e9-130">Além disso, o arquivo popup.js não contém qualquer JavaScript incompatível com o Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="595e9-130">In addition, the popup.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="595e9-131">Por esses dois motivos, esse suplemento não transcompila o popup.js.</span><span class="sxs-lookup"><span data-stu-id="595e9-131">For these two reasons, this add-in does not transpile the popup.js file at all.</span></span> 


## <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="595e9-132">Abra a caixa de diálogo do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="595e9-132">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="595e9-133">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="595e9-133">Open the file index.html.</span></span>
2. <span data-ttu-id="595e9-134">Abaixo do `div` que contém o botão `freeze-header`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="595e9-134">Below the `div` that contains the `freeze-header` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="open-dialog">Open Dialog</button>          
    </div>
    ```

3. <span data-ttu-id="595e9-135">A caixa de diálogo solicitará que o usuário insira um nome e passará o nome de usuário para o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="595e9-135">The dialog will prompt the user to enter a name and pass the user's name to the task pane.</span></span> <span data-ttu-id="595e9-136">O painel de tarefas o exibirá em um rótulo.</span><span class="sxs-lookup"><span data-stu-id="595e9-136">The task pane will display it in a label.</span></span> <span data-ttu-id="595e9-137">Imediatamente abaixo do `div` que você adicionou, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="595e9-137">Immediately below the `div` that you just added, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <label id="user-name"></label>            
    </div>
    ```

4. <span data-ttu-id="595e9-138">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="595e9-138">Open the app.js file.</span></span>

5. <span data-ttu-id="595e9-139">Abaixo da linha que atribui um identificador de clique ao botão `freeze-header`, adicione o seguinte código.</span><span class="sxs-lookup"><span data-stu-id="595e9-139">Below the line that assigns a click handler to the `freeze-header` button, add the following code.</span></span> <span data-ttu-id="595e9-140">Você criará o método `openDialog` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="595e9-140">You'll create the `openDialog` method in a later step.</span></span>

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. <span data-ttu-id="595e9-p112">Abaixo da função `freezeHeader`, adicione a declaração seguinte. Essa variável é usada para armazenar um objeto no contexto de execução da página pai que atua como um intermediador no contexto de execução da página da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="595e9-p112">Below the `freezeHeader` function add the following declaration. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    let dialog = null;
    ```

7. <span data-ttu-id="595e9-143">Abaixo da declaração de `dialog`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="595e9-143">Below the declaration of `dialog`, add the following function.</span></span> <span data-ttu-id="595e9-144">É importante observar o que esse código *não* contém: não há nenhuma chamada de `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="595e9-144">The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`.</span></span> <span data-ttu-id="595e9-145">Isso ocorre porque a API para abrir uma caixa de diálogo é compartilhada com todos os hosts do Office, portanto, ela faz parte da API de Office JavaScript Common, não da API específica do Excel.</span><span class="sxs-lookup"><span data-stu-id="595e9-145">This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Shared API that opens a dialog
    }
    ``` 

8. <span data-ttu-id="595e9-p114">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="595e9-p114">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="595e9-148">O método`displayDialogAsync` abre uma caixa de diálogo no centro da tela.</span><span class="sxs-lookup"><span data-stu-id="595e9-148">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>
   - <span data-ttu-id="595e9-149">O primeiro parâmetro é a URL da página a ser aberta.</span><span class="sxs-lookup"><span data-stu-id="595e9-149">The first parameter is the URL of the page to open.</span></span>
   - <span data-ttu-id="595e9-p115">O segundo parâmetro passa opções. `height` e `width` são porcentagens do tamanho da janela do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="595e9-p115">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span> 
   
    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},
        
        // TODO2: Add callback parameter.
    );
    ``` 

## <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="595e9-152">Processar a mensagem da caixa de diálogo e depois fechá-la</span><span class="sxs-lookup"><span data-stu-id="595e9-152">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="595e9-p116">Continue no arquivo app.js e substitua `TODO2` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="595e9-p116">Continue in the app.js file, and replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="595e9-155">O retorno de chamada é executado logo após a caixa de diálogo ser aberta com êxito e antes de o usuário executar qualquer ação nela.</span><span class="sxs-lookup"><span data-stu-id="595e9-155">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>
   - <span data-ttu-id="595e9-156">O `result.value` é o objeto que funciona como um tipo de intermediário entre contextos execução das páginas de pai e de caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="595e9-156">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>
   - <span data-ttu-id="595e9-157">A função `processMessage` será criada em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="595e9-157">The `processMessage` function will be created in a later step.</span></span> <span data-ttu-id="595e9-158">Esse identificador processará os valores que sejam enviados da página da caixa de diálogo com chamadas da função `messageParent`.</span><span class="sxs-lookup"><span data-stu-id="595e9-158">This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="595e9-159">Abaixo da função `openDialog`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="595e9-159">Below the `openDialog` function, add the following function.</span></span>

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="595e9-160">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="595e9-160">Test the add-in</span></span>

1. <span data-ttu-id="595e9-161">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl + C duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="595e9-161">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="595e9-162">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="595e9-162">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="595e9-163">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="595e9-163">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="595e9-164">Para fazer isso, interrompa o processo do servidor para obter uma solicitação para inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="595e9-164">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="595e9-165">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="595e9-165">After the build, you restart the server.</span></span> <span data-ttu-id="595e9-166">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="595e9-166">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="595e9-167">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="595e9-167">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="595e9-168">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="595e9-168">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="595e9-169">Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="595e9-169">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
6. <span data-ttu-id="595e9-170">Escolha o botão **Abrir Caixa de Diálogo** no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="595e9-170">Choose the **Open Dialog** button in the task pane.</span></span> 
7. <span data-ttu-id="595e9-171">Quando a caixa de diálogo estiver aberta, arraste-a e redimensione-a.</span><span class="sxs-lookup"><span data-stu-id="595e9-171">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="595e9-172">Observe que você pode interagir com a planilha e pressionar outros botões no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="595e9-172">Note that you can interact with the worksheet and press other buttons on the taskpane.</span></span> <span data-ttu-id="595e9-173">No entanto, não é possível iniciar uma segunda caixa de diálogo na mesma página do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="595e9-173">But you cannot launch a second dialog from the same task pane page.</span></span>
8. <span data-ttu-id="595e9-174">Na caixa de diálogo, digite um nome e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="595e9-174">In the dialog, enter a name and choose **OK**.</span></span> <span data-ttu-id="595e9-175">O nome aparecerá no painel de tarefas e a caixa de diálogo será fechada.</span><span class="sxs-lookup"><span data-stu-id="595e9-175">The name appears on the task pane and the dialog closes.</span></span>
9. <span data-ttu-id="595e9-176">Opcionalmente, comente a linha `dialog.close();` na função `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="595e9-176">Optionally, comment out the line `dialog.close();` in the `processMessage` function.</span></span> <span data-ttu-id="595e9-177">Em seguida, repita as etapas desta seção.</span><span class="sxs-lookup"><span data-stu-id="595e9-177">Then repeat the steps of this section.</span></span> <span data-ttu-id="595e9-178">A caixa de diálogo permanece aberta e você pode alterar o nome.</span><span class="sxs-lookup"><span data-stu-id="595e9-178">The dialog stays open and you can change the name.</span></span> <span data-ttu-id="595e9-179">É possível fechá-la manualmente pressionando o botão **X** no canto superior direito.</span><span class="sxs-lookup"><span data-stu-id="595e9-179">You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Tutorial do Excel - Caixa de diálogo](../images/excel-tutorial-dialog-open.png)

