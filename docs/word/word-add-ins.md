# <a name="build-your-first-word-add-in"></a>Compilar seu primeiro suplemento do Word

_Aplica-se a: Word 2016, Word para iPad, Word para Mac_

Um suplemento do Word é executado no Word e pode interagir com o conteúdo do documento usando a API JavaScript para Word, que faz parte do modelo de programação dos Suplementos do Office para estender aplicativos do Office. Neste modelo de programação do suplemento, você pode usar a plataforma e o idioma de sua preferência para criar o aplicativo Web que hospeda sua extensão no Word e usar o [manifesto](../overview/add-in-manifests.md) do suplemento para definir suas configurações e recursos.

Neste artigo, você passará pelo processo de criar um suplemento do Word usando o jQuery e a API JavaScript para Word. 

> **Observação**: para desenvolver um suplemento para o Word 2013, será preciso usar a [API Javascript para Office]( https://dev.office.com/docs/add-ins/word/word-add-ins-programming-overview#javascript-apis-for-word) compartilhada. Saiba mais sobre as plataformas e as diferentes APIs que estão disponíveis em [Disponibilidade de host e plataforma para Suplementos do Office](https://dev.office.com/add-in-availability). 

## <a name="create-the-web-app"></a>Criar o aplicativo Web 

1. Crie uma pasta na sua unidade local e nomeie-a como **BoilerplateAddin**. Esse é o local em que você criará os arquivos para seu aplicativo.

2. Na pasta do aplicativo, crie um arquivo chamado **home.html** para especificar o HTML que será renderizado no painel de tarefas do suplemento. Este suplemento exibirá três botões, e quando qualquer um dos botões for escolhido, o texto clichê será adicionado ao documento. Adicione o código a seguir a **home.html** e salve o arquivo.

    ```html
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Boilerplate text app</title>
        <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
        <script src="home.js" type="text/javascript"></script>
        </head>
        <body>
            <div>
                <h1>Welcome</h1>
            </div>
            <div>
                <p>This sample shows how to add boilerplate text to a document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <button id="proverb">Add Chinese proverb</button>
            </div>
            <h3><div id="supportedVersion"/></h3>
        </body>
    </html>
    ```

3. Na pasta do aplicativo, crie um arquivo chamado **home.js** para especificar o script jQuery para o suplemento. Esse script contém códigos de inicialização além do código que faz alterações no documento do Word inserindo texto no documento quando um botão é escolhido. Adicione o código a seguir a **home.js** e salve o arquivo.

    ```javascript
    (function () {
        "use strict";

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

## <a name="create-the-manifest-file"></a>Criar o arquivo de manifesto.

1. Na pasta do aplicativo, crie um arquivo chamado **BoilerplateManifest.xml** para definir as configurações e os recursos do suplemento. Adicione o código a seguir a esse arquivo. 

    ```xml
    <?xml version="1.0" encoding="UTF-8"?>
        <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xsi:type="TaskPaneApp">
            <Id>2b88100c-656e-4bab-9f1e-f6731d86e464</Id>
            <Version>1.0.0.0</Version>
            <ProviderName>Microsoft</ProviderName>
            <DefaultLocale>en-US</DefaultLocale>
            <DisplayName DefaultValue="Boilerplate content" />
            <Description DefaultValue="Insert boilerplate content into a Word document." />
            <Hosts>
                <Host Name="Document"/>
            </Hosts>
            <DefaultSettings>
                <SourceLocation DefaultValue="\\MyShare\boilerplate\home.html" />
            </DefaultSettings>
            <Permissions>ReadWriteDocument</Permissions>
        </OfficeApp>
    ```

2. Gere um GUID usando um gerador online de sua preferência. Em seguida, substitua o valor do elemento **Id** mostrado na etapa anterior por esse GUID.

3. Salve o arquivo de manifesto.

## <a name="deploy-the-web-app-and-update-the-manifest"></a>Implantar o aplicativo Web e atualizar o manifesto

1. Implante o aplicativo Web (por exemplo, o conteúdo da sua pasta de aplicativo) no servidor Web de sua escolha.

2. Na sua pasta local do aplicativo, abra o arquivo de manifesto (**BoilerplateManifest.xml**). Edite o valor do atributo no elemento **SourceLocation** para especificar o local do arquivo **home.html** no servidor Web e salve o arquivo.

## <a name="try-it-out"></a>Experimente

1. Siga as instruções para a plataforma que você usará para executar o suplemento e fazer sideload do suplemento no Word.

    - Windows: [Fazer sideload dos Suplementos do Office para teste no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Word Online: [Fazer sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad e Mac: [Fazer sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. No painel à direita, escolha qualquer um dos botões para adicionar o texto clichê ao documento.

![Imagem do aplicativo Word com o suplemento de texto clichê carregado](../images/boilerplateAddin.png)

## <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com êxito um suplemento do Word usando o jQuery! Em seguida, saiba mais sobre os [principais conceitos](word-add-ins-programming-overview.md) de criação de suplementos do Word.

## <a name="additional-resources"></a>Recursos adicionais

* [Visão geral dos suplementos do Word](word-add-ins-programming-overview.md)
* [Explorar trechos com o Script Lab](https://store.office.com/en-001/app.aspx?assetid=WA104380862&ui=en-US&rs=en-001&ad=US&appredirect=false)
* [Exemplos de código do suplemento do Word](http://dev.office.com/code-samples#?filters=word,office%20add-ins)
* [Referências da API JavaScript do Word](http://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)