# <a name="build-your-first-project-add-in"></a>Criar o primeiro suplemento do Project

Neste artigo, você passará pelo processo de criar um suplemento do Project usando o jQuery e a API JavaScript para Office.

## <a name="prerequisites"></a>Pré-requisitos

- [Node.js](https://nodejs.org)

- Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a>Criar o suplemento

1. Crie uma pasta na sua unidade local e nomeie-a como `my-project-addin`. Esse é o local em que você criará os arquivos para seu suplemento.

2. Navegue até a nova pasta.

    ```bash
    cd my-project-addin
    ```

3. Use o gerador Yeoman para criar um projeto de suplemento do Project. Execute o comando a seguir e responda aos prompts da seguinte forma:

    ```bash
    yo office
    ```

    - **Gostaria de criar uma nova subpasta para o seu projeto?** `No`
    - **Como deseja nomear seu suplemento?:** `My Office Add-in`
    - **Para qual aplicativo cliente do Office você deseja suporte?** `Project`
    - **Gostaria de criar um novo suplemento?:** `Yes`
    - **Gostaria de usar o TypeScript?** `No`
    - **Escolha a estrutura:** `Jquery`

    O gerador perguntará se você deseja abrir **resource.html**. Não é necessário abri-lo para este tutorial, mas fique à vontade em fazer isso se tiver curiosidade. Escolha Sim ou Não para concluir o assistente e deixar o gerador fazer seu trabalho.

    ![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-project-jquery.png)

## <a name="update-the-code"></a>Atualizar o código

1. No editor de código, abra **index.html** na raiz do projeto. Esse arquivo contém o HTML que será renderizado no painel de tarefas do suplemento.

2. Substitua o elemento `<header>` dentro do elemento `<body>` com a marcação a seguir.

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

3. Substitua o elemento `<main>` dentro do elemento `<body>` com a marcação a seguir e salve o arquivo.

    ```html
    <div id="content-main">
        <div class="padding">
            <p>Select a task and then choose the buttons below and observe the output in the <b>Results</b> textbox.</p>
            <h3>Try it out</h3>
            <button class="ms-Button" id="get-task-guid">Get Task GUID</button>
            <br/><br/>
            <button class="ms-Button" id="get-task">Get Task data</button>
            <br/>
            <h4>Results:</h4>
            <textarea id="result" rows="6" cols="25"></textarea>
        </div>
    </div>
    ```

4. Abra o arquivo **app.js** para especificar o script do suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.

    ```js
    'use strict';

    (function () {

        var taskGuid;

        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#get-task-guid').click(getTaskGUID);
                $('#get-task').click(getTask);
            });
        };

        function getTaskGUID() {
            Office.context.document.getSelectedTaskAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    result.value = "Task GUID: " + asyncResult.value;
                    taskGuid = asyncResult.value;
                }
                else {
                    console.log(asyncResult.error.message);
                }
            });
        }

        function getTask() {
            if (taskGuid != undefined) {
                Office.context.document.getTaskAsync(
                    taskGuid,
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            var taskInfo = asyncResult.value;
                            var taskOutput = "Task name: " + taskInfo.taskName +
                                            "\nGUID: " + taskGuid +
                                            "\nWSS Id: " + taskInfo.wssTaskId +
                                            "\nResource names: " + taskInfo.resourceNames;
                            result.value = taskOutput;
                        } else {
                            console.log(asyncResult.error.message);
                        }
                    }
                );
            } else {
                result.value = 'Task GUID not valid:\n' + taskGuid;
            } 
        }
    })();
    ```

4. Abra o arquivo **app.css** na raiz do projeto para especificar os estilos personalizados do suplemento. Substitua todo o conteúdo pelo que está a seguir e salve o arquivo.

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

## <a name="update-the-manifest"></a>Atualizar o manifesto

1. Abra o arquivo **my-office-add-in-manifest.xml** para definir as configurações e os recursos do suplemento.

2. O elemento `ProviderName` tem um valor de espaço reservado. Substitua-o com seu nome.

3. O atributo `DefaultValue` do elemento `Description` tem um espaço reservado. Substitua-o com **um suplemento do painel de tarefas do Project**.

4. Salve o arquivo.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Project"/>
    ...
    ```

## <a name="start-the-dev-server"></a>Iniciar o servidor de desenvolvimento

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a>Experimente

1. No Project, crie um projeto simples que tenha pelo menos uma tarefa.

2. Siga as instruções para a plataforma que você usará para executar o suplemento e para fazer o sideload do suplemento no Project.

    - Windows: [Realizar o sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Project Online: [Realizar o sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad e Mac: [Fazer sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

3. No Project, selecione uma tarefa.

    ![Uma captura de tela de um plano de projeto no Project com uma tarefa selecionada](../images/project_quickstart_addin_1.png)

4. No painel de tarefas, escolha o botão **Obter GUID de tarefas** para gravar a GUID de tarefas na caixa de texto **Resultados**.

    ![Captura de tela de um plano de projeto no Project com uma tarefa selecionada e a GUID de tarefas gravada na caixa de texto no painel de tarefas](../images/project_quickstart_addin_2.png)

5. No painel de tarefas, escolha o botão **Obter dados da tarefa** para gravar várias propriedades da tarefa selecionada na caixa de texto **Resultados**.

    ![Captura de tela de um plano de projeto no Project com uma tarefa selecionada e várias propriedades de tarefas gravadas na caixa de texto do painel de tarefas](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com êxito um suplemento do Project! Em seguida, saiba mais sobre os recursos de um suplemento do Project e explore os cenários comuns.

> [!div class="nextstepaction"]
> [Suplementos do Project](../project/project-add-ins.md)
