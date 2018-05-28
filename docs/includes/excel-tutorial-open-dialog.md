Nesta etapa final do tutorial, voc? abre uma caixa de di?logo no suplemento, passa uma mensagem do processo de caixa de di?logo para o processo de painel de tarefas e fecha a caixa de di?logo. As caixas de di?logo do Suplemento do Office s?o *n?o modais*: o usu?rio pode continuar a interagir com o documento no aplicativo do Office do host e com a p?gina host no painel de tarefas.

> [!NOTE]
> Esta p?gina descreve uma etapa individual do tutorial de suplemento do Excel. Se voc? chegou aqui por meio dos resultados de mecanismos de pesquisa ou via outro link direto, acesse a p?gina de Introdu??o do [tutorial de suplemento do Excel](../tutorials/excel-tutorial.yml) para come??-lo do in?cio.

## <a name="create-the-dialog-page"></a>Crie a p?gina da caixa de di?logo

1. Abra o projeto em seu editor de c?digo.
2. Crie um arquivo chamado popup.html na raiz do projeto (onde se encontra index.html).
3. Adicione a marca??o a seguir em popup.html. Observa??o:
   - A p?gina tem um `<input>` onde o usu?rio insere seu nome e um bot?o que enviar? o nome para a p?gina no painel de tarefas onde ele ser? exibido.
   - A marca??o carrega um script chamado popup.js que voc? criar? em uma etapa posterior.
   - Ela tamb?m carrega uma biblioteca Office.JS e jQuery porque elas ser?o usadas em popup.js.

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

4. Crie um arquivo chamado popup.js na raiz do projeto.
5. Adicione o c?digo a seguir ao popup.js. Observa??o:
   - *Todas as p?ginas que chamam APIs na biblioteca Office.JS devem atribuir uma fun??o ? propriedade `Office.initialize`.* Se nenhuma inicializa??o for necess?ria, a fun??o poder? ter um corpo vazio, mas a propriedade n?o deve ser deixada indefinida, atribu?da a nulo ou a um valor que n?o seja uma fun??o. Por exemplo, veja o arquivo app.js na raiz do projeto. O c?digo que cria a tarefa deve ser executado antes de qualquer chamada para Office.JS; por isso, a tarefa se encontra em um arquivo de script que ? carregado pela p?gina, como neste caso.
   - A fun??o jQuery `ready` ? chamada dentro do m?todo `initialize`. ? uma regra quase universal que o c?digo de carregamento, inicializa??o ou bootstrapping de outras bibliotecas JavaScript deva estar dentro da fun??o `Office.initialize`.

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

6. Substitua `TODO1` pelo c?digo a seguir. Voc? criar? a fun??o `sendStringToParentPage` na pr?xima etapa.

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. Substitua `TODO2` pelo c?digo a seguir. O m?todo `messageParent` passa seu par?metro para a p?gina pai, neste caso, a p?gina no painel de tarefas. O par?metro pode ser um booliano ou uma cadeia de caracteres, que inclui tudo o que pode ser serializado como uma cadeia de caracteres, como XML ou JSON. 

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. Salve o arquivo.

   > [!NOTE]
   > O arquivo popup.html e o arquivo popup.js carregado s?o executados em um processo do Internet Explorer completamente separado de painel de tarefas do suplemento. Se o popup.js foi transcompilado no mesmo arquivo bundle.js que o arquivo app.js, o suplemento precisar? carregar duas c?pias do arquivo bundle.js, o que anule o prop?sito do agrupamento. Al?m disso, o arquivo popup.js n?o cont?m qualquer JavaScript incompat?vel com o Internet Explorer. Por esses dois motivos, esse suplemento n?o transcompila o popup.js. 


## <a name="open-the-dialog-from-the-task-pane"></a>Abra a caixa de di?logo do painel de tarefas

1. Abra o arquivo index.html.
2. Abaixo do `div` que cont?m o bot?o `freeze-header`, adicione a marca??o a seguir:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="open-dialog">Open Dialog</button>          
    </div>
    ```

3. A caixa de di?logo solicitar? que o usu?rio insira um nome e passar? o nome de usu?rio para o painel de tarefas. O painel de tarefas o exibir? em um r?tulo. Imediatamente abaixo do `div` que voc? adicionou, adicione a marca??o a seguir:

    ```html
    <div class="padding">            
        <label id="user-name"></label>            
    </div>
    ```

4. Abra o arquivo app.js.

5. Abaixo da linha que atribui um identificador de clique ao bot?o `freeze-header`, adicione o seguinte c?digo. Voc? criar? o m?todo `openDialog` em uma etapa posterior.

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. Abaixo da fun??o `freezeHeader`, adicione a declara??o seguinte. Essa vari?vel ? usada para armazenar um objeto no contexto de execu??o da p?gina pai que atua como um intermediador no contexto de execu??o da p?gina da caixa de di?logo.

    ```js
    let dialog = null;
    ```

7. Abaixo da declara??o de `dialog`, adicione a fun??o a seguir. ? importante observar o que esse c?digo *n?o* cont?m: n?o h? nenhuma chamada de `Excel.run`. Isso ocorre porque a API para abrir uma caixa de di?logo ? compartilhada com todos os hosts do Office, portanto, ela faz parte da API de Office JavaScript Common, n?o da API espec?fica do Excel.

    ```js
    function openDialog() {
        // TODO1: Call the Office Shared API that opens a dialog
    }
    ``` 

8. Substitua `TODO1` pelo c?digo a seguir. Observa??o:
   - O m?todo`displayDialogAsync` abre uma caixa de di?logo no centro da tela.
   - O primeiro par?metro ? a URL da p?gina a ser aberta.
   - O segundo par?metro passa op??es. `height` e `width` s?o porcentagens do tamanho da janela do aplicativo do Office. 
   
    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},
        
        // TODO2: Add callback parameter.
    );
    ``` 

## <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a>Processar a mensagem da caixa de di?logo e depois fech?-la

1. Continue no arquivo app.js e substitua `TODO2` pelo c?digo a seguir. Observa??o:
   - O retorno de chamada ? executado logo ap?s a caixa de di?logo ser aberta com ?xito e antes de o usu?rio executar qualquer a??o nela.
   - O `result.value` ? o objeto que funciona como um tipo de intermedi?rio entre contextos execu??o das p?ginas de pai e de caixa de di?logo.
   - A fun??o `processMessage` ser? criada em uma etapa posterior. Esse identificador processar? os valores que sejam enviados da p?gina da caixa de di?logo com chamadas da fun??o `messageParent`.

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. Abaixo da fun??o `openDialog`, adicione a fun??o a seguir.

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

## <a name="test-the-add-in"></a>Testar o suplemento

1. Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl + C duas vezes para interromper a execu??o do servidor Web. Caso contr?rio, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue at? a pasta **Iniciar** do projeto.

     > [!NOTE]
     > Embora o servidor de sincroniza??o do navegador recarregue o suplemento no painel de tarefas sempre que voc? fizer uma altera??o em algum arquivo, incluindo o arquivo app.js, ele n?o transcompila o JavaScript, portanto, ? necess?rio repetir o comando de compila??o para que as altera??es em app.js as entrem em vigor. Para fazer isso, interrompa o processo do servidor para obter uma solicita??o para inserir o comando de compila??o. Ap?s a compila??o, reinicie o servidor. As pr?ximas etapas executam esse processo.

1. Execute o comando `npm run build` para transcompilar seu c?digo-fonte ES6 para uma vers?o anterior do JavaScript com suporte no Internet Explorer (que ? usada em segundo plano pelo Excel para executar os suplementos do Excel).
2. Execute o comando `npm start` para iniciar um servidor Web em um host local.
4. Feche o painel de tarefas para recarreg?-lo e, no menu **In?cio**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.
6. Escolha o bot?o **Abrir Caixa de Di?logo** no painel de tarefas. 
7. Quando a caixa de di?logo estiver aberta, arraste-a e redimensione-a. Observe que voc? pode interagir com a planilha e pressionar outros bot?es no painel de tarefas. No entanto, n?o ? poss?vel iniciar uma segunda caixa de di?logo na mesma p?gina do painel de tarefas.
8. Na caixa de di?logo, digite um nome e escolha **OK**. O nome aparecer? no painel de tarefas e a caixa de di?logo ser? fechada.
9. Opcionalmente, comente a linha `dialog.close();` na fun??o `processMessage`. Em seguida, repita as etapas desta se??o. A caixa de di?logo permanece aberta e voc? pode alterar o nome. ? poss?vel fech?-la manualmente pressionando o bot?o **X** no canto superior direito.

    ![Tutorial do Excel - Caixa de di?logo](../images/excel-tutorial-dialog-open.png)

