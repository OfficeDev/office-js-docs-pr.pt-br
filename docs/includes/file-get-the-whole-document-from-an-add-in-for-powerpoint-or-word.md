Você pode criar um suplemento do Office para oferecer o envio ou a publicação de um documento do Word 2013 ou do PowerPoint 2013 para um local remoto com um único clique. Este artigo demonstra como criar um suplemento de painel de tarefas simples para o PowerPoint 2013 que obtém todas as apresentações como um objeto de dados e envia esses dados para um servidor Web por meio de uma solicitação HTTP.

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a>Pré-requisitos para a criação de um suplemento para o PowerPoint ou Word

Este artigo pressupõe que você esteja usando um editor de texto para criar o suplemento de painel de tarefas do PowerPoint ou Word. Para criar o suplemento do painel de tarefas, você deve criar os arquivos a seguir.

- Em uma pasta de rede compartilhada ou em um servidor Web, você precisa dos arquivos a seguir.

  - Um arquivo HTML (GetDoc_App.html) que contém a interface do usuário e links para os arquivos JavaScript (incluindo arquivos office.js e arquivos de .js específicos do aplicativo) e arquivos CSS (Folha de Estilos em Cascata).

  - Um arquivo JavaScript (GetDoc_App.js) para conter a lógica de programação do suplemento.

  - Um arquivo CSS (Program.css) para conter os estilos e formatação do suplemento.

- Um arquivo de manifesto XML (GetDoc_App.xml) para o suplemento, disponível em uma pasta de rede compartilhada ou catálogo de suplementos. O arquivo de manifesto deve apontar para o local do arquivo HTML mencionado anteriormente.

Você também pode criar um suplemento para PowerPoint usando o [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio) ou o gerador [Yeoman para suplementos do Office](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator) ou para o Word usando o [Visual Studio](../quickstarts/word-quickstart.md?tabs=visualstudio) ou o gerador [Yeoman para suplementos do Office](../quickstarts/word-quickstart.md?tabs=yeomangenerator).

### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a>Conceitos fundamentais para a criação de um suplemento de painel de tarefas

Antes de começar a criar esse suplemento do PowerPoint ou Word, você deve estar familiarizado com a criação de suplementos do Office e com o trabalho com solicitações HTTP. Este artigo não discute como decodificar texto codificado em Base64 de uma solicitação HTTP em um servidor Web.

## <a name="create-the-manifest-for-the-add-in"></a>Criar o manifesto para o suplemento

O arquivo de manifesto XML para o suplemento do PowerPoint fornece informações importantes sobre o suplemento: quais aplicativos podem hospedá-lo, o local do arquivo HTML, o título e a descrição do suplemento e muitas outras características.

1. Em um editor de texto, adicione o seguinte código ao arquivo do manifesto.

    ```xml  
    <?xml version="1.0" encoding="utf-8" ?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xsi:type="TaskPaneApp">
        <Id>[Replace_With_Your_GUID]</Id>
        <Version>1.0</Version>
        <ProviderName>[Provider Name]</ProviderName>
        <DefaultLocale>EN-US</DefaultLocale>
        <DisplayName DefaultValue="Get Doc add-in" />
        <Description DefaultValue="My get PowerPoint or Word document add-in." />
        <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
        <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
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

O exemplo de código a seguir mostra o manipulador de eventos `Office.initialize` para o evento juntamente com uma função auxiliar, `updateStatus`para gravar no div de status.

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
    statusInfo[0].innerHTML += message + "<br/>";
}
```

Quando você escolhe o **botão** Enviar na interface do usuário, o suplemento `sendFile` chama a função, que contém uma chamada para o [método Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) . O `getFileAsync` método usa o padrão assíncrono, semelhante a outros métodos na API JavaScript para Office. Ele tem um parâmetro obrigatório, _fileType_, e dois parâmetros opcionais,  _options_ e _callback_.

O _parâmetro fileType_ espera uma das três constantes da enumeração [FileType](/javascript/api/office/office.filetype): `Office.FileType.Compressed` ("compactada **"),Office.FileType.PDF** ("pdf") ou **Office. FileType.Text** ("text"). O suporte de tipo de arquivo atual para cada plataforma está listado nos comentários [Document.getFileType](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) . Quando você passa Compactado para o parâmetro _fileType_, `getFileAsync` o método retorna o documento como um arquivo de apresentação do PowerPoint 2013 (.pptx) ou arquivo de documento do *Word 2013 (*.docx) criando uma cópia temporária do arquivo no computador local.

O `getFileAsync` método retorna uma referência ao arquivo como um [objeto File](/javascript/api/office/office.file) . O `File` objeto expõe quatro membros: a propriedade [size](/javascript/api/office/office.file#office-office-file-size-member) , [a propriedade sliceCount](/javascript/api/office/office.file#office-office-file-slicecount-member) , o método [getSliceAsync](/javascript/api/office/office.file#office-office-file-getsliceasync-member(1)) e [o método closeAsync](/javascript/api/office/office.file#office-office-file-closeasync-member(1)) . A `size` propriedade retorna o número de bytes no arquivo. Retorna `sliceCount` o número de [objetos Slice](/javascript/api/office/office.slice) (discutidos mais adiante neste artigo) no arquivo.

Use o código a seguir para obter o PowerPoint ou o documento do Word `File` `Document.getFileAsync` como um objeto usando o método e, em seguida, faz uma chamada para a função definida `getSlice` localmente. Observe que o objeto `File` , uma variável de contador e o número total de fatias no arquivo são passados `getSlice` na chamada para um objeto anônimo.

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

A função local `getSlice` faz uma chamada para o `File.getSliceAsync` método para recuperar uma fatia do `File` objeto. O `getSliceAsync` método retorna um `Slice` objeto da coleção de fatias. Ele tem dois parâmetros obrigatórios, _sliceIndex_ e _callback_. O parâmetro _sliceIndex_ usa um número inteiro como um indexador na coleção de fatias. Assim como outras funções na API JavaScript para Office, `getSliceAsync` o método também usa uma função de retorno de chamada como um parâmetro para manipular os resultados da chamada de método.
O ion `getSlice` faz uma chamada para o **método File.getSliceAsync** para recuperar uma fatia do **objeto** File. O método **getSliceAsync** retorna um objeto **Slice** do conjunto de fatias. Ele tem dois parâmetros obrigatórios, _sliceIndex_ e _callback_. O parâmetro _sliceIndex_ usa um número inteiro como um indexador na coleção de fatias. Assim como outras funções na API JavaScript Office, o método **getSliceAsync** também usa uma função de retorno de chamada como um parâmetro para manipular os resultados da chamada de método.

O `Slice` objeto fornece acesso aos dados contidos no arquivo. A menos que especificado de outra forma _no_ parâmetro de opções do `getFileAsync` método, o `Slice` objeto tem 4 MB de tamanho. O `Slice` objeto expõe três propriedades: [tamanho](/javascript/api/office/office.slice#office-office-slice-size-member), [dados](/javascript/api/office/office.slice#office-office-slice-data-member) e [índice](/javascript/api/office/office.slice#office-office-slice-index-member). A `size` propriedade obtém o tamanho, em bytes, da fatia. A `index` propriedade obtém um inteiro que representa a posição da fatia na coleção de fatias.

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

A `Slice.data` propriedade retorna os dados brutos do arquivo como uma matriz de bytes. Se os dados forem no formato de texto (ou seja, XML ou texto sem formatação), a fatia contém o texto não processado. Se você passar em **Office. FileType.Compressed para** o _parâmetro fileType_ `Document.getFileAsync`de , a fatia contém os dados binários do arquivo como uma matriz de bytes. No caso de um arquivo do PowerPoint ou do Word, as fatias contêm matrizes de bytes.

Você deve implementar sua própria função (ou usar uma biblioteca disponível) para converter dados da matriz de bytes em uma cadeia de caracteres com codificação por Base64. Para saber mais sobre a codificação por Base64 com JavaScript, confira [Codificação e decodificação por Base64](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).

Depois de converter os dados para Base64, você pode transmiti-los para um servidor Web de diversas maneiras, incluindo como corpo de uma solicitação HTTP POST.

Adicione o seguinte código para enviar uma fatia para um serviço Web.

> [!NOTE]
> Esse código envia um PowerPoint ou arquivo do Word para o servidor Web em várias fatias. O servidor Web ou serviço deve acrescentar cada fatia individual em um único arquivo e, em seguida, salvá-la como um arquivo .pptx ou .docx, antes que você possa executar qualquer manipulação nele.

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

Como o nome sugere, o `File.closeAsync` método fecha a conexão com o documento e libera recursos. Embora o lixo de área restrita dos Suplementos do Office colete referências fora do escopo para arquivos, ainda é uma prática recomendada fechar explicitamente os arquivos depois que o código não precisar mais deles. O `closeAsync` método tem um único parâmetro, _retorno_ de chamada, que especifica a função a ser chamada na conclusão da chamada.

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