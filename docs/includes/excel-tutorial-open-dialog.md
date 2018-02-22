Nesta etapa final do tutorial, você abre uma caixa de diálogo no suplemento, passa uma mensagem do processo de caixa de diálogo para o processo de painel de tarefas e fecha a caixa de diálogo. As caixas de diálogo do Suplemento do Office são *não modais*: o usuário pode continuar a interagir com o documento no aplicativo do Office do host e com a página host no painel de tarefas.

## <a name="create-the-dialog-page"></a>Crie a página da caixa de diálogo

1. Abra o projeto em seu editor de código.
2. Crie um arquivo na raiz do projeto (onde index.html se encontra) chamado popup.html.
3. Adicione a marcação a seguir em popup.html. Observação:
   - A página tem um `<input>` onde o usuário insere o nome dele e um botão que enviará o nome para a página no painel de tarefas onde ele será exibido.
   - A marcação carrega um script chamado popup.js que você criará em uma etapa posterior.
   - Ela também carrega uma biblioteca Office.JS e jQuery porque elas serão usadas em popup.js.

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

4. Crie um arquivo na raiz do projeto chamado o popup.js.
5. Adicione o código a seguir a popup.js. Observação:
   - *Todas as páginas que chamam APIs na biblioteca Office.JS devem atribuir uma função à propriedade `Office.initialize`.* Se nenhuma inicialização for necessária, a função poderá ter um corpo vazio, mas a propriedade não deve ser deixada indefinida, atribuída a nulo ou a um valor que não seja uma função. Por exemplo, veja o arquivo app.js na raiz do projeto. O código que cria a tarefa deve ser executado antes de qualquer chamada para Office.JS; por isso, a tarefa se encontra em um arquivo de script que é carregado pela página, como neste caso.
   - A função `ready` do jQuery é chamada dentro do método `initialize`. É uma regra quase universal que o código de carregamento, inicialização ou bootstrapping de outras bibliotecas JavaScript devem estar dentro da função `Office.initialize`.

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

6. Substitua `TODO1` pelo código a seguir. Você criará a função `sendStringToParentPage` na próxima etapa.

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. Substitua `TODO2` pelo código a seguir. O método `messageParent` passa seu parâmetro para a página pai, neste caso, a página no painel de tarefas. O parâmetro pode ser um booliano ou uma cadeia de caracteres, que inclui tudo o que pode ser serializado como uma cadeia de caracteres, como XML ou JSON. 

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. Salve o arquivo.

   > [!NOTE]
   > O arquivo popup.html e o arquivo popup.js carregado são executados em um processo do Internet Explorer completamente separado de painel de tarefas do suplemento. Se o popup.js foi transcompilado no mesmo arquivo bundle.js que o arquivo app.js, o suplemento precisará carregar duas cópias do arquivo bundle.js, o que anule o propósito do agrupamento. Além disso, o arquivo popup.js não contém qualquer JavaScript incompatível com o Internet Explorer. Por esses dois motivos, esse suplemento não transcompila o popup.js. 


## <a name="open-the-dialog-from-the-task-pane"></a>Abra a caixa de diálogo do painel de tarefas

1. Abra o arquivo index.html.
2. Abaixo do `div` que contém o botão `freeze-header`, adicione a marcação a seguir:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="open-dialog">Open Dialog</button>          
    </div>
    ```

3. A caixa de diálogo solicitará que o usuário insira um nome e passará o nome de usuário para o painel de tarefas. O painel de tarefas o exibirá em um rótulo. Imediatamente abaixo do `div` que você adicionou, adicione a marcação a seguir:

    ```html
    <div class="padding">            
        <label id="user-name"></label>            
    </div>
    ```

4. Abra o arquivo app.js.

5. Abaixo da linha que atribui um identificador de clique ao botão `freeze-header`, adicione o seguinte código. Você criará o método `openDialog` em uma etapa posterior.

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. Abaixo da função `freezeHeader`, adicione a declaração seguir. Essa variável é usada para armazenar um objeto no contexto de execução da página pai que atua como um intermediador no contexto de execução da página da caixa de diálogo.

    ```js
    let dialog = null;
    ```

7. Abaixo da declaração de `dialog`, adicione a função seguir. É importante observar o que esse código *não* contém: não há nenhuma chamada de `Excel.run`. Isso ocorre porque a API para abrir uma caixa de diálogo é compartilhada com todos os hosts do Office, portanto, ela faz parte da API de Office JavaScript Common, não da API específica do Excel.

    ```js
    function openDialog() {
        // TODO1: Call the Office Shared API that opens a dialog
    }
    ``` 

8. Substitua `TODO1` pelo código a seguir. Observação:
   - O método`displayDialogAsync` abre uma caixa de diálogo no centro da tela.
   - O primeiro parâmetro é a URL da página a ser aberta.
   - O segundo parâmetro passa opções. `height` e `width` são porcentagens do tamanho da janela do aplicativo do Office. 
   
    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},
        
        // TODO2: Add callback parameter.
    );
    ``` 

## <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a>Processe a mensagem da caixa de diálogo e feche a caixa de diálogo

1. Continue no arquivo app.js e substitua `TODO2` pelo código a seguir. Observação:
   - O retorno de chamada é executado imediatamente depois que a caixa de diálogo é aberta com êxito e antes de usuário executar a ação na caixa de diálogo.
   - O `result.value` é o objeto que funciona como um tipo de intermediário entre contextos execução das páginas de pai e de caixa de diálogo.
   - A função `processMessage` será criada em uma etapa posterior. Esse identificador processará os valores que sejam enviados da página da caixa de diálogo com chamadas da função `messageParent`.

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. Abaixo da função `openDialog`, adicione a função a seguir.

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

## <a name="test-the-add-in"></a>Teste o suplemento

1. Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiver aberto, insira Ctrl + C duas vezes para interromper a execução do servidor Web. Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.

     > [!NOTE]
     > Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor. Para fazer isso, interrompa o processo do servidor para obter uma solicitação para inserir o comando de compilação. Após a compilação, reinicie o servidor. As próximas etapas executam esse processo.

1. Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).
2. Execute o comando `npm start` para iniciar um servidor Web em um host local.
4. Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.
6. Escolha o botão **Abrir Caixa de Diálogo** no painel de tarefas. 
7. Quando a caixa de diálogo estiver aberta, arraste-a e redimensione-a. Observe que você pode interagir com a planilha e pressionar outros botões no painel de tarefas. No entanto, não é possível iniciar uma segunda caixa de diálogo na mesma página do painel de tarefas.
8. Na caixa de diálogo, digite um nome e escolha **OK**. O nome aparecerá no painel de tarefas e a caixa de diálogo será fechada.
9. Opcionalmente, comente a linha `dialog.close();` na função `processMessage`. Em seguida, repita as etapas desta seção. A caixa de diálogo permanece aberta e você pode alterar o nome. É possível fechá-la manualmente pressionando o botão **X** no canto superior direito.

    ![Tutorial do Excel - Caixa de diálogo](../images/excel-tutorial-dialog-open.png)

