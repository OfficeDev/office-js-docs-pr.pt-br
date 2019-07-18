# <a name="get-the-whole-document-from-an-add-in-for-powerpoint-or-word"></a>Obter todo o documento por meio de um suplemento para PowerPoint ou Word

Você pode criar um suplemento do Office para oferecer o envio ou a publicação de um documento do Word 2013 ou do PowerPoint 2013 para um local remoto com um único clique. Este artigo demonstra como criar um suplemento de painel de tarefas simples para o PowerPoint 2013 que obtém todas as apresentações como um objeto de dados e envia esses dados para um servidor Web por meio de uma solicitação HTTP.

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a>Pré-requisitos para a criação de um suplemento para o PowerPoint ou Word

Este artigo pressupõe que você esteja usando um editor de texto para criar o suplemento de painel de tarefas do PowerPoint ou Word. Para criar o suplemento de painel de tarefas, você deve criar os seguintes arquivos:

- Em uma pasta de rede compartilhada ou em um servidor Web, você precisará dos seguintes arquivos:

    - Um arquivo HTML (GetDoc_App.html) contendo a interface do usuário mais links para os arquivos de JavaScript (incluindo arquivos office.js e .js específico do host) e arquivos de Folha de Estilos em Cascata (CSS).

    - Um arquivo JavaScript (GetDoc_App.js) para conter a lógica de programação do suplemento.

    - Um arquivo CSS (Program.css) para conter os estilos e formatação do suplemento.

- Um arquivo de manifesto XML (GetDoc_App.xml) para o suplemento, disponível em uma pasta de rede compartilhada ou catálogo de suplementos. O arquivo de manifesto deve apontar para o local do arquivo HTML mencionado anteriormente.

Você também pode criar um suplemento para o PowerPoint usando o [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio) ou o [gerador Yeoman para suplementos do Office](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator) ou para o Word usando o [Visual Studio](../quickstarts/word-quickstart.md?tabs=visualstudio) ou o [gerador do Yeoman para suplementos do Office](../quickstarts/word-quickstart.md?tabs=yeomangenerator).

### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a>Conceitos fundamentais para a criação de um suplemento de painel de tarefas

Antes de começar a criar esse suplemento do PowerPoint ou Word, você deve estar familiarizado com a criação de suplementos do Office e com o trabalho com solicitações HTTP. Este artigo não aborda como decodificar textos com codificação Base64 de uma solicitação HTTP em um servidor Web. 

## <a name="create-the-manifest-for-the-add-in"></a>Criar o manifesto para o suplemento

O arquivo de manifesto XML para o suplemento do PowerPoint fornece informações importantes sobre o suplemento: quais aplicativos podem hospedá-lo, o local do arquivo HTML, o título e a descrição do suplemento e muitas outras características.

1. Em um editor de texto, adicione o seguinte código ao arquivo do manifesto.

    ```xml  
    <?xml version="1.0" encoding="utf-8" ?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
    xsi:type="TaskPaneApp">
        <Id>[Replace_With_Your_GUID]</Id>
        <Version>1.0</Version>
        <ProviderName>[Provider Name]</ProviderName>
        <DefaultLocale>EN-US</DefaultLocale>
        <DisplayName DefaultValue="Get Doc add-in" />
        <Description DefaultValue="My get PowerPoint or Word document add-in." />
        <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
        <Hosts>
        <Host Name="Document" />
        <Host Name="Presentation" />
        </Hosts>
        <DefaultSettings>
        <SourceLocation DefaultValue="[Network location of app]/GetDoc_App.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

2. Salve o arquivo como GetDoc_App.xml, usando a codificação UTF-8, em um local de rede ou um catálogo de suplemento.

## <a name="create-the-user-interface-for-the-add-in"></a>Criar a interface de usuário para o suplemento

Para a interface de usuário do suplemento, você pode usar HTML escrito diretamente no arquivo GetDoc_App.html. A lógica de programação e a funcionalidade do suplemento devem estar contidos em um arquivo JavaScript (por exemplo, GetDoc_App.js).

Use o procedimento a seguir para criar uma interface de usuário simples para o suplemento incluindo um cabeçalho e um único botão.

1. Em um novo arquivo no editor de texto, adicione o seguinte HTML.

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
            <title>Publish presentation</title>
            <link rel="stylesheet" type="text/css" href="Program.css" />
            <script src="https://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js" type="text/javascript"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
            <script src="GetDoc_App.js"></script>
        </head>
        <body>
        <form>
            <h1>Publish presentation</h1>
            <br />
            <div><input id='submit' type="button" value="Submit" /></div>
            <br />
            <div><h2>Status</h2> 
                <div id="status"></div>
            </div>
        </form>
        </body>
    </html>
    ```

2. Salve o arquivo como GetDoc_App.html, usando a codificação UTF-8, em um local de rede ou um servidor Web.

    > [!NOTE]
    > Certifique-se de que as marcas **head** do suplemento contenham uma marca **script** com um link válido para o arquivo office.js. 

    Vamos usar alguns CSS para dar ao suplemento uma aparência simples, porém moderna e profissional. Use os seguintes CSS para definir o estilo do suplemento.

3. Em um novo arquivo no editor de texto, adicione o seguinte CSS.

    ```css  
    body
    {
        font-family: "Segoe UI Light","Segoe UI",Tahoma,sans-serif;
    }
    h1,h2
    {
        text-decoration-color:#4ec724;
    }
    input [type="submit"], input[type="button"]
    {
        height:24px;
        padding-left:1em;
        padding-right:1em;
        background-color:white;
        border:1px solid grey;
        border-color: #dedfe0 #b9b9b9 #b9b9b9 #dedfe0;
        cursor:pointer;
    }
    ```

4. Salve o arquivo como Program.css, utilizando a codificação UTF-8, no local de rede ou servidor Web em que o arquivo GetDoc_App.html está localizado.

## <a name="add-the-javascript-to-get-the-document"></a>Adicionar o JavaScript para obter o documento

No código para o suplemento, um manipulador para o evento [Office.initialize](/javascript/api/office) adiciona um manipulador para o evento de clique do botão **Enviar** no formulário e informa aos usuários que o suplemento está pronto.

O exemplo de código a seguir mostra o manipulador de eventos do evento **Office.initialize** juntamente com a função auxiliar, `updateStatus`, para escrever na div de status.

```js
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {

        // Execute sendFile when submit is clicked
        $('#submit').click(function () {
            sendFile();
        });

        // Update status
        updateStatus("Ready to send file.");
    });
}

// Create a function for writing to the status div.
function updateStatus(message) {
    var statusInfo = $('#status');
    statusInfo.innerHTML += message + "<br/>";
}
```

Quando você escolhe o botão **Enviar** na interface do usuário, o suplemento chama a função `sendFile`, que contém uma chamada para o método [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-). O método **getFileAsync** usa o padrão assíncrono, semelhante a outros métodos na API JavaScript para Office. Ele tem um parâmetro obrigatório, _fileType_, e dois parâmetros opcionais,  _options_ e _callback_. 

O parâmetro _fileType_ espera uma das três constantes da enumeração [FileType](/javascript/api/office/office.filetype): **Office.FileType.Compressed** ("compactado"), **Office.FileType.PDF** ("PDF"), ou **Office.FileType.Text** ("texto"). O PowerPoint só suporta **Compressed** como argumento; o Word suporta todos os três. Quando você transmite **Compressed** para o parâmetro _fileType_, o método **getFileAsync** retorna o documento como um arquivo de apresentação do PowerPoint 2013 (*.pptx) ou arquivo de documento do Word 2013 (*.docx) criando uma cópia temporária do arquivo no computador local.

O método **getFileAsync** retorna uma referência para o arquivo como um objeto [File](/javascript/api/office/office.file). O objeto **File** expõe quatro membros: a propriedade [size](/javascript/api/office/office.file#size), a propriedade [sliceCount](/javascript/api/office/office.file#slicecount), o método [getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) e o método [closeAsync](/javascript/api/office/office.file#closeasync-callback-). A propriedade **size** retorna o número de bytes no arquivo. A propriedade **sliceCount** retorna o número de objetos [Slice](/javascript/api/office/office.slice) (será discutido posteriormente neste artigo) no arquivo.

Use o código a seguir para obter o documento do PowerPoint ou Word como um objeto **File** usando o método **Document.getFileAsync** e, em seguida, faça uma chamada para a função `getSlice` definida localmente. Observe que o objeto **File**, uma variável de contador e o número total de fatias no arquivo são transmitidos na chamada para `getSlice` em um objeto anônimo.

```js
// Get all of the content from a PowerPoint or Word document in 100-KB chunks of text.
function sendFile() {
    Office.context.document.getFileAsync("compressed",
        { sliceSize: 100000 },
        function (result) {

            if (result.status == Office.AsyncResultStatus.Succeeded) {

                // Get the File object from the result.
                var myFile = result.value;
                var state = {
                    file: myFile,
                    counter: 0,
                    sliceCount: myFile.sliceCount
                };

                updateStatus("Getting file of " + myFile.size + " bytes");
                getSlice(state);
            }
            else {
                updateStatus(result.status);
            }
        });
}
```

A função local `getSlice` faz uma chamada para o método **File.getSliceAsync** para recuperar uma fatia do objeto **File**. O método **getSliceAsync** retorna um objeto **Slice** do conjunto de fatias. Ele tem dois parâmetros obrigatórios, _sliceIndex_ e _callback_. O parâmetro _sliceIndex_ usa um número inteiro como um indexador na coleção de fatias. Como outras funções na API JavaScript para Office, o método **getSliceAsync** também usa uma função de retorno de chamada como um parâmetro para manipular os resultados da chamada do método.

O objeto **Slice** dá acesso aos dados do arquivo. A menos que seja especificado de outra forma no parâmetro _options_ do método **getFileAsync**, o objeto **Slice** tem 4 MB de tamanho. O objeto **Slice** expõe três propriedades: [size](/javascript/api/office/office.slice#size), [data](/javascript/api/office/office.slice#data) e [index](/javascript/api/office/office.slice#index). A propriedade **size** obtém o tamanho, em bytes, da fatia. A propriedade **index** obtém um número inteiro que representa a posição da fatia na coleção de fatias.

```js
// Get a slice from the file and then call sendSlice.
function getSlice(state) {
    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {
            updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
            sendSlice(result.value, state);
        }
        else {
            updateStatus(result.status);
        }
    });
}
```

A propriedade **Slice.data** retorna os dados brutos do arquivo como uma matriz de bytes. Se os dados forem no formato de texto (ou seja, XML ou texto sem formatação), a fatia contém o texto não processado. Se você transmitir o **Office.FileType.Compressed** para o parâmetro _fileType_ de **Document.getFileAsync**, a fatia contém os dados binários do arquivo como uma matriz de byte. No caso de um arquivo do PowerPoint ou do Word, as fatias contêm matrizes de bytes.

Você deve implementar sua própria função (ou usar uma biblioteca disponível) para converter dados da matriz de bytes em uma cadeia de caracteres com codificação por Base64. Para saber mais sobre a codificação por Base64 com JavaScript, confira [Codificação e decodificação por Base64](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).

Depois de converter os dados para Base64, você pode transmiti-los para um servidor Web de diversas maneiras, incluindo como corpo de uma solicitação HTTP POST.

Adicione o seguinte código para enviar uma fatia para um serviço Web.

> [!NOTE]
> Este código envia um arquivo do PowerPoint ou Word para o servidor Web em várias fatias. O servidor Web ou serviço deve acrescentar cada fatia individual em um único arquivo e, em seguida, salvá-lo como um arquivo. pptx ou. docx, antes de poder executar qualquer manipulação nele.

```js
function sendSlice(slice, state) {
    var data = slice.data;

    // If the slice contains data, create an HTTP request.
    if (data) {

        // Encode the slice data, a byte array, as a Base64 string.
        // NOTE: The implementation of myEncodeBase64(input) function isn't
        // included with this example. For information about Base64 encoding with
        // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
        var fileData = myEncodeBase64(data);

        // Create a new HTTP request. You need to send the request
        // to a webpage that can receive a post.
        var request = new XMLHttpRequest();

        // Create a handler function to update the status
        // when the request has been sent.
        request.onreadystatechange = function () {
            if (request.readyState == 4) {

                updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;

                if (state.counter < state.sliceCount) {
                    getSlice(state);
                }
                else {
                    closeFile(state);
                }
            }
        }

        request.open("POST", "[Your receiving page or service]");
        request.setRequestHeader("Slice-Number", slice.index);

        // Send the file as the body of an HTTP POST
        // request to the web server.
        request.send(fileData);
    }
}
```

Como o nome sugere, o método **File.closeAsync** fecha a conexão com o documento e libera os recursos. Embora o lixo de área restrita dos Suplementos do Office colete referências fora do escopo para arquivos, ainda é uma prática recomendada fechar explicitamente os arquivos depois que o código não precisar mais deles. O método **closeAsync** tem um único parâmetro, _callback_, que especifica a função a ser chamada na conclusão da chamada.

```js
function closeFile(state) {
    // Close the file when you're done with it.
    state.file.closeAsync(function (result) {

        // If the result returns as a success, the
        // file has been successfully closed.
        if (result.status == "succeeded") {
            updateStatus("File closed.");
        }
        else {
            updateStatus("File couldn't be closed.");
        }
    });
}
```