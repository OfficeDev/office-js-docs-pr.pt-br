<span data-ttu-id="5014f-p101">Você pode criar um suplemento do Office para oferecer o envio ou a publicação de um documento do Word 2013 ou do PowerPoint 2013 para um local remoto com um único clique. Este artigo demonstra como criar um suplemento de painel de tarefas simples para o PowerPoint 2013 que obtém todas as apresentações como um objeto de dados e envia esses dados para um servidor Web por meio de uma solicitação HTTP.</span><span class="sxs-lookup"><span data-stu-id="5014f-p101">You can create an Office Add-in to provide one-click sending or publishing of a Word 2013 or PowerPoint 2013 document to a remote location. This article demonstrates how to build a simple task pane add-in for PowerPoint 2013 that gets all of the presentation as a data object and sends that data to a web server via an HTTP request.</span></span>

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a><span data-ttu-id="5014f-103">Pré-requisitos para a criação de um suplemento para o PowerPoint ou Word</span><span class="sxs-lookup"><span data-stu-id="5014f-103">Prerequisites for creating an add-in for PowerPoint or Word</span></span>

<span data-ttu-id="5014f-p102">Este artigo pressupõe que você esteja usando um editor de texto para criar o suplemento de painel de tarefas do PowerPoint ou Word. Para criar o suplemento de painel de tarefas, você deve criar os seguintes arquivos:</span><span class="sxs-lookup"><span data-stu-id="5014f-p102">This article assumes that you are using a text editor to create the task pane add-in for PowerPoint or Word. To create the task pane add-in, you must create the following files:</span></span>

- <span data-ttu-id="5014f-106">Em uma pasta de rede compartilhada ou em um servidor Web, você precisará dos seguintes arquivos:</span><span class="sxs-lookup"><span data-stu-id="5014f-106">On a shared network folder or on a web server, you need the following files:</span></span>

    - <span data-ttu-id="5014f-107">Um arquivo HTML (GetDoc_App.html) contendo a interface do usuário mais links para os arquivos de JavaScript (incluindo arquivos office.js e .js específico do host) e arquivos de Folha de Estilos em Cascata (CSS).</span><span class="sxs-lookup"><span data-stu-id="5014f-107">An HTML file (GetDoc_App.html) that contains the user interface plus links to the JavaScript files (including office.js and host-specific .js files) and Cascading Style Sheet (CSS) files.</span></span>

    - <span data-ttu-id="5014f-108">Um arquivo JavaScript (GetDoc_App.js) para conter a lógica de programação do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5014f-108">A JavaScript file (GetDoc_App.js) to contain the programming logic of the add-in.</span></span>

    - <span data-ttu-id="5014f-109">Um arquivo CSS (Program.css) para conter os estilos e formatação do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5014f-109">A CSS file (Program.css) to contain the styles and formatting for the add-in.</span></span>

- <span data-ttu-id="5014f-p103">Um arquivo de manifesto XML (GetDoc_App.xml) para o suplemento, disponível em uma pasta de rede compartilhada ou catálogo de suplementos. O arquivo de manifesto deve apontar para o local do arquivo HTML mencionado anteriormente.</span><span class="sxs-lookup"><span data-stu-id="5014f-p103">An XML manifest file (GetDoc_App.xml) for the add-in, available on a shared network folder or add-in catalog. The manifest file must point to the location of the HTML file mentioned previously.</span></span>

<span data-ttu-id="5014f-112">Você também pode criar um suplemento para o PowerPoint usando o [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio) ou o [gerador Yeoman para suplementos do Office](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator) ou para o Word usando o [Visual Studio](../quickstarts/word-quickstart.md?tabs=visualstudio) ou o [gerador do Yeoman para suplementos do Office](../quickstarts/word-quickstart.md?tabs=yeomangenerator).</span><span class="sxs-lookup"><span data-stu-id="5014f-112">You can also create an add-in for PowerPoint by using [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio) or the [Yeoman generator for Office Add-ins](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator) or for Word by using [Visual Studio](../quickstarts/word-quickstart.md?tabs=visualstudio) or [Yeoman generator for Office Add-ins](../quickstarts/word-quickstart.md?tabs=yeomangenerator).</span></span>

### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a><span data-ttu-id="5014f-113">Conceitos fundamentais para a criação de um suplemento de painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="5014f-113">Core concepts to know for creating a task pane add-in</span></span>

<span data-ttu-id="5014f-p104">Antes de começar a criar esse suplemento do PowerPoint ou Word, você deve estar familiarizado com a criação de suplementos do Office e com o trabalho com solicitações HTTP. Este artigo não aborda como decodificar textos com codificação Base64 de uma solicitação HTTP em um servidor Web.</span><span class="sxs-lookup"><span data-stu-id="5014f-p104">Before you begin creating this add-in for PowerPoint or Word, you should be familiar with building Office Add-ins and working with HTTP requests. This article does not discuss how to decode Base64-encoded text from an HTTP request on a web server.</span></span> 

## <a name="create-the-manifest-for-the-add-in"></a><span data-ttu-id="5014f-116">Criar o manifesto para o suplemento</span><span class="sxs-lookup"><span data-stu-id="5014f-116">Create the manifest for the add-in</span></span>

<span data-ttu-id="5014f-117">O arquivo de manifesto XML para o suplemento do PowerPoint fornece informações importantes sobre o suplemento: quais aplicativos podem hospedá-lo, o local do arquivo HTML, o título e a descrição do suplemento e muitas outras características.</span><span class="sxs-lookup"><span data-stu-id="5014f-117">The XML manifest file for the add-in for PowerPoint provides important information about the add-in: what applications can host it, the location of the HTML file, the add-in title and description, and many other characteristics.</span></span>

1. <span data-ttu-id="5014f-118">Em um editor de texto, adicione o seguinte código ao arquivo do manifesto.</span><span class="sxs-lookup"><span data-stu-id="5014f-118">In a text editor, add the following code to the manifest file.</span></span>

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

2. <span data-ttu-id="5014f-119">Salve o arquivo como GetDoc_App.xml, usando a codificação UTF-8, em um local de rede ou um catálogo de suplemento.</span><span class="sxs-lookup"><span data-stu-id="5014f-119">Save the file as GetDoc_App.xml using UTF-8 encoding to a network location or to an add-in catalog.</span></span>

## <a name="create-the-user-interface-for-the-add-in"></a><span data-ttu-id="5014f-120">Criar a interface de usuário para o suplemento</span><span class="sxs-lookup"><span data-stu-id="5014f-120">Create the user interface for the add-in</span></span>

<span data-ttu-id="5014f-p105">Para a interface de usuário do suplemento, você pode usar HTML escrito diretamente no arquivo GetDoc_App.html. A lógica de programação e a funcionalidade do suplemento devem estar contidos em um arquivo JavaScript (por exemplo, GetDoc_App.js).</span><span class="sxs-lookup"><span data-stu-id="5014f-p105">For the user interface of the add-in, you can use HTML, written directly into the GetDoc_App.html file. The programming logic and functionality of the add-in must be contained in a JavaScript file (for example, GetDoc_App.js).</span></span>

<span data-ttu-id="5014f-123">Use o procedimento a seguir para criar uma interface de usuário simples para o suplemento incluindo um cabeçalho e um único botão.</span><span class="sxs-lookup"><span data-stu-id="5014f-123">Use the following procedure to create a simple user interface for the add-in that includes a heading and a single button.</span></span>

1. <span data-ttu-id="5014f-124">Em um novo arquivo no editor de texto, adicione o seguinte HTML.</span><span class="sxs-lookup"><span data-stu-id="5014f-124">In a new file in the text editor, add the following HTML.</span></span>

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

2. <span data-ttu-id="5014f-125">Salve o arquivo como GetDoc_App.html, usando a codificação UTF-8, em um local de rede ou um servidor Web.</span><span class="sxs-lookup"><span data-stu-id="5014f-125">Save the file as GetDoc_App.html using UTF-8 encoding to a network location or to a web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5014f-126">Certifique-se de que as marcas **head** do suplemento contenham uma marca **script** com um link válido para o arquivo office.js.</span><span class="sxs-lookup"><span data-stu-id="5014f-126">Be sure that the **head** tags of the add-in contains a **script** tag with a valid link to the office.js file.</span></span> 

    <span data-ttu-id="5014f-p106">Vamos usar alguns CSS para dar ao suplemento uma aparência simples, porém moderna e profissional. Use os seguintes CSS para definir o estilo do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5014f-p106">We'll use some CSS to give the add-in a simple, yet modern and professional appearance. Use the following CSS to define the style of the add-in.</span></span>

3. <span data-ttu-id="5014f-129">Em um novo arquivo no editor de texto, adicione o seguinte CSS.</span><span class="sxs-lookup"><span data-stu-id="5014f-129">In a new file in the text editor, add the following CSS.</span></span>

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

4. <span data-ttu-id="5014f-130">Salve o arquivo como Program.css, utilizando a codificação UTF-8, no local de rede ou servidor Web em que o arquivo GetDoc_App.html está localizado.</span><span class="sxs-lookup"><span data-stu-id="5014f-130">Save the file as Program.css using UTF-8 encoding to the network location or to the web server where the GetDoc_App.html file is located.</span></span>

## <a name="add-the-javascript-to-get-the-document"></a><span data-ttu-id="5014f-131">Adicionar o JavaScript para obter o documento</span><span class="sxs-lookup"><span data-stu-id="5014f-131">Add the JavaScript to get the document</span></span>

<span data-ttu-id="5014f-132">No código para o suplemento, um manipulador para o evento [Office.initialize](/javascript/api/office) adiciona um manipulador para o evento de clique do botão **Enviar** no formulário e informa aos usuários que o suplemento está pronto.</span><span class="sxs-lookup"><span data-stu-id="5014f-132">In the code for the add-in, a handler to the [Office.initialize](/javascript/api/office) event adds a handler to the click event of the **Submit** button on the form and informs the user that the add-in is ready.</span></span>

<span data-ttu-id="5014f-133">O exemplo de código a seguir mostra o manipulador de eventos do evento **Office.initialize** juntamente com a função auxiliar, `updateStatus`, para escrever na div de status.</span><span class="sxs-lookup"><span data-stu-id="5014f-133">The following code example shows the event handler for the  **Office.initialize** event along with a helper function, `updateStatus`, for writing to the status div.</span></span>

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

<span data-ttu-id="5014f-p107">Quando você escolhe o botão **Enviar** na interface do usuário, o suplemento chama a função `sendFile`, que contém uma chamada para o método [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-). O método **getFileAsync** usa o padrão assíncrono, semelhante a outros métodos na API JavaScript para Office. Ele tem um parâmetro obrigatório, _fileType_, e dois parâmetros opcionais,  _options_ e _callback_.</span><span class="sxs-lookup"><span data-stu-id="5014f-p107">When you choose the  **Submit** button in the UI, the add-in calls the `sendFile` function, which contains a call to the [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) method. The **getFileAsync** method uses the asynchronous pattern, similar to other methods in the JavaScript API for Office. It has one required parameter, _fileType_, and two optional parameters,  _options_ and _callback_.</span></span> 

<span data-ttu-id="5014f-p108">O parâmetro _fileType_ espera uma das três constantes da enumeração [FileType](/javascript/api/office/office.filetype): **Office.FileType.Compressed** ("compactado"), **Office.FileType.PDF** ("PDF"), ou **Office.FileType.Text** ("texto"). O PowerPoint só suporta **Compressed** como argumento; o Word suporta todos os três. Quando você transmite **Compressed** para o parâmetro _fileType_, o método **getFileAsync** retorna o documento como um arquivo de apresentação do PowerPoint 2013 (*.pptx) ou arquivo de documento do Word 2013 (*.docx) criando uma cópia temporária do arquivo no computador local.</span><span class="sxs-lookup"><span data-stu-id="5014f-p108">The  _fileType_ parameter expects one of three constants from the [FileType](/javascript/api/office/office.filetype) enumeration: **Office.FileType.Compressed** ("compressed"), **Office.FileType.PDF** ("pdf"), or **Office.FileType.Text** ("text"). PowerPoint supports only **Compressed** as an argument; Word supports all three. When you pass in **Compressed** for the _fileType_ parameter, the **getFileAsync** method returns the document as a PowerPoint 2013 presentation file (*.pptx) or Word 2013 document file (*.docx) by creating a temporary copy of the file on the local computer.</span></span>

<span data-ttu-id="5014f-p109">O método **getFileAsync** retorna uma referência para o arquivo como um objeto [File](/javascript/api/office/office.file). O objeto **File** expõe quatro membros: a propriedade [size](/javascript/api/office/office.file#size), a propriedade [sliceCount](/javascript/api/office/office.file#slicecount), o método [getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) e o método [closeAsync](/javascript/api/office/office.file#closeasync-callback-). A propriedade **size** retorna o número de bytes no arquivo. A propriedade **sliceCount** retorna o número de objetos [Slice](/javascript/api/office/office.slice) (será discutido posteriormente neste artigo) no arquivo.</span><span class="sxs-lookup"><span data-stu-id="5014f-p109">The  **getFileAsync** method returns a reference to the file as a [File](/javascript/api/office/office.file) object. The **File** object exposes four members: the [size](/javascript/api/office/office.file#size) property, [sliceCount](/javascript/api/office/office.file#slicecount) property, [getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) method, and [closeAsync](/javascript/api/office/office.file#closeasync-callback-) method. The **size** property returns the number of bytes in the file. The **sliceCount** returns the number of [Slice](/javascript/api/office/office.slice) objects (discussed later in this article) in the file.</span></span>

<span data-ttu-id="5014f-p110">Use o código a seguir para obter o documento do PowerPoint ou Word como um objeto **File** usando o método **Document.getFileAsync** e, em seguida, faça uma chamada para a função `getSlice` definida localmente. Observe que o objeto **File**, uma variável de contador e o número total de fatias no arquivo são transmitidos na chamada para `getSlice` em um objeto anônimo.</span><span class="sxs-lookup"><span data-stu-id="5014f-p110">Use the following code to get the PowerPoint or Word document as a  **File** object using the **Document.getFileAsync** method and then makes a call to the locally defined `getSlice` function. Note that the **File** object, a counter variable, and the total number of slices in the file are passed along in the call to `getSlice` in an anonymous object.</span></span>

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

<span data-ttu-id="5014f-p111">A função local `getSlice` faz uma chamada para o método **File.getSliceAsync** para recuperar uma fatia do objeto **File**. O método **getSliceAsync** retorna um objeto **Slice** do conjunto de fatias. Ele tem dois parâmetros obrigatórios, _sliceIndex_ e _callback_. O parâmetro _sliceIndex_ usa um número inteiro como um indexador na coleção de fatias. Como outras funções na API JavaScript para Office, o método **getSliceAsync** também usa uma função de retorno de chamada como um parâmetro para manipular os resultados da chamada do método.</span><span class="sxs-lookup"><span data-stu-id="5014f-p111">The local function  `getSlice` makes a call to the **File.getSliceAsync** method to retrieve a slice from the **File** object. The **getSliceAsync** method returns a **Slice** object from the collection of slices. It has two required parameters, _sliceIndex_ and _callback_. The  _sliceIndex_ parameter takes an integer as an indexer into the collection of slices. Like other functions in the JavaScript API for Office, the **getSliceAsync** method also takes a callback function as a parameter to handle the results from the method call.</span></span>

<span data-ttu-id="5014f-151">O objeto **Slice** dá acesso aos dados do arquivo.</span><span class="sxs-lookup"><span data-stu-id="5014f-151">The **Slice** object gives you access to the data contained in the file.</span></span> <span data-ttu-id="5014f-152">A menos que seja especificado de outra forma no parâmetro _options_ do método **getFileAsync**, o objeto **Slice** tem 4 MB de tamanho.</span><span class="sxs-lookup"><span data-stu-id="5014f-152">Unless otherwise specified in the _options_ parameter of the **getFileAsync** method, the **Slice** object is 4 MB in size.</span></span> <span data-ttu-id="5014f-153">O objeto **Slice** expõe três propriedades: [size](/javascript/api/office/office.slice#size), [data](/javascript/api/office/office.slice#data) e [index](/javascript/api/office/office.slice#index).</span><span class="sxs-lookup"><span data-stu-id="5014f-153">The **Slice** object exposes three properties: [size](/javascript/api/office/office.slice#size), [data](/javascript/api/office/office.slice#data), and [index](/javascript/api/office/office.slice#index).</span></span> <span data-ttu-id="5014f-154">A propriedade **size** obtém o tamanho, em bytes, da fatia.</span><span class="sxs-lookup"><span data-stu-id="5014f-154">The **size** property gets the size, in bytes, of the slice.</span></span> <span data-ttu-id="5014f-155">A propriedade **index** obtém um número inteiro que representa a posição da fatia na coleção de fatias.</span><span class="sxs-lookup"><span data-stu-id="5014f-155">The **index** property gets an integer that represents the slice's position in the collection of slices.</span></span>

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

<span data-ttu-id="5014f-p113">A propriedade **Slice.data** retorna os dados brutos do arquivo como uma matriz de bytes. Se os dados forem no formato de texto (ou seja, XML ou texto sem formatação), a fatia contém o texto não processado. Se você transmitir o **Office.FileType.Compressed** para o parâmetro _fileType_ de **Document.getFileAsync**, a fatia contém os dados binários do arquivo como uma matriz de byte. No caso de um arquivo do PowerPoint ou do Word, as fatias contêm matrizes de bytes.</span><span class="sxs-lookup"><span data-stu-id="5014f-p113">The  **Slice.data** property returns the raw data of the file as a byte array. If the data is in text format (that is, XML or plain text), the slice contains the raw text. If you pass in **Office.FileType.Compressed** for the _fileType_ parameter of **Document.getFileAsync**, the slice contains the binary data of the file as a byte array. In the case of a PowerPoint or Word file, the slices contain byte arrays.</span></span>

<span data-ttu-id="5014f-p114">Você deve implementar sua própria função (ou usar uma biblioteca disponível) para converter dados da matriz de bytes em uma cadeia de caracteres com codificação por Base64. Para saber mais sobre a codificação por Base64 com JavaScript, confira [Codificação e decodificação por Base64](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).</span><span class="sxs-lookup"><span data-stu-id="5014f-p114">You must implement your own function (or use an available library) to convert byte array data to a Base64-encoded string. For information about Base64 encoding with JavaScript, see [Base64 encoding and decoding](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).</span></span>

<span data-ttu-id="5014f-162">Depois de converter os dados para Base64, você pode transmiti-los para um servidor Web de diversas maneiras, incluindo como corpo de uma solicitação HTTP POST.</span><span class="sxs-lookup"><span data-stu-id="5014f-162">Once you have converted the data to Base64, you can then transmit it to a web server in several ways -- including as the body of an HTTP POST request.</span></span>

<span data-ttu-id="5014f-163">Adicione o seguinte código para enviar uma fatia para um serviço Web.</span><span class="sxs-lookup"><span data-stu-id="5014f-163">Add the following code to send a slice to a web service.</span></span>

> [!NOTE]
> <span data-ttu-id="5014f-164">Este código envia um arquivo do PowerPoint ou Word para o servidor Web em várias fatias.</span><span class="sxs-lookup"><span data-stu-id="5014f-164">This code sends a PowerPoint or Word file to the web server in multiple slices.</span></span> <span data-ttu-id="5014f-165">O servidor Web ou serviço deve acrescentar cada fatia individual em um único arquivo e, em seguida, salvá-lo como um arquivo. pptx ou. docx, antes de poder executar qualquer manipulação nele.</span><span class="sxs-lookup"><span data-stu-id="5014f-165">The web server or service must append each individual slice into a single file, and then save it as a .pptx or .docx file, before you can perform any manipulations on it.</span></span>

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

<span data-ttu-id="5014f-p116">Como o nome sugere, o método **File.closeAsync** fecha a conexão com o documento e libera os recursos. Embora o lixo de área restrita dos Suplementos do Office colete referências fora do escopo para arquivos, ainda é uma prática recomendada fechar explicitamente os arquivos depois que o código não precisar mais deles. O método **closeAsync** tem um único parâmetro, _callback_, que especifica a função a ser chamada na conclusão da chamada.</span><span class="sxs-lookup"><span data-stu-id="5014f-p116">As the name implies, the  **File.closeAsync** method closes the connection to the document and frees up resources. Although the Office Add-ins sandbox garbage collects out-of-scope references to files, it is still a best practice to explicitly close files once your code is done with them. The **closeAsync** method has a single parameter, _callback_, that specifies the function to call on the completion of the call.</span></span>

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