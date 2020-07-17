---
title: Suplementos do PowerPoint
description: Aprenda a usar os suplementos do PowerPoint para criar soluções atraentes para apresentações em plataformas como Windows, iPad, Mac e em um navegador.
ms.date: 06/29/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 314b441f3d4b6d2188ed630fe2b254aec42a86bc
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006448"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="6883d-103">Suplementos do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="6883d-103">PowerPoint add-ins</span></span>

<span data-ttu-id="6883d-104">Você pode usar suplementos do PowerPoint na criação de soluções envolventes para as apresentações de seus usuários em todas as plataformas, incluindo Windows, iPad, Mac e em um navegador.</span><span class="sxs-lookup"><span data-stu-id="6883d-104">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iPad, Mac, and in a browser.</span></span> <span data-ttu-id="6883d-105">Você pode criar dois tipos de comandos de suplementos do PowerPoint:</span><span class="sxs-lookup"><span data-stu-id="6883d-105">You can create two types of PowerPoint add-ins:</span></span>

- <span data-ttu-id="6883d-p102">Use **suplementos de conteúdo** para adicionar conteúdo dinâmico do HTML5 às suas apresentações. Por exemplo, confira o suplemento [LucidChart Diagrams for PowerPoint](https://appsource.microsoft.com/product/office/wa104380117) que pode ser usado para inserir um diagrama interativo do LucidChart para seu conjunto.</span><span class="sxs-lookup"><span data-stu-id="6883d-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://appsource.microsoft.com/product/office/wa104380117) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>

- <span data-ttu-id="6883d-108">Use **suplementos do painel de tarefas** para exibir as informações de referência ou inserir dados na apresentação através de um serviço.</span><span class="sxs-lookup"><span data-stu-id="6883d-108">Use **task pane add-ins** to bring in reference information or insert data into the presentation via a service.</span></span> <span data-ttu-id="6883d-109">Por exemplo, consulte o suplemento [Pexels Free Stock Photos](https://appsource.microsoft.com/product/office/wa104379997), que você pode usar para adicionar fotos profissionais à sua apresentação.</span><span class="sxs-lookup"><span data-stu-id="6883d-109">For example, see the [Pexels - Free Stock Photos](https://appsource.microsoft.com/product/office/wa104379997) add-in, which you can use to add professional photos to your presentation.</span></span>

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="6883d-110">Cenários de suplemento do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="6883d-110">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="6883d-111">Os exemplos de código no artigo mostram algumas tarefas básicas para desenvolver suplementos para o PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="6883d-111">The code examples in this article demonstrate some basic tasks for developing add-ins for PowerPoint.</span></span> <span data-ttu-id="6883d-112">Além disso, observe o seguinte:</span><span class="sxs-lookup"><span data-stu-id="6883d-112">Please note the following:</span></span>

- <span data-ttu-id="6883d-113">Para exibir as informações, esses exemplos dependem da função `app.showNotification`, incluída em modelos de projeto de Suplementos do Office do Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="6883d-113">To display information, these examples use the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates.</span></span> <span data-ttu-id="6883d-114">Se você não estiver usando o Visual Studio para desenvolver seu suplemento, será necessário substituir a função `showNotification` por seu próprio código.</span><span class="sxs-lookup"><span data-stu-id="6883d-114">If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code.</span></span>

- <span data-ttu-id="6883d-115">Vários desses exemplos também usam um objeto `Globals` que é declarado fora do âmbito destas funções como:   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span><span class="sxs-lookup"><span data-stu-id="6883d-115">Several of these examples also use a `Globals` object that is declared beyond the scope of these functions as:   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

- <span data-ttu-id="6883d-116">Para usar esses exemplos, seu suplemento de projeto deverá [referenciar a biblioteca do Office.js 1.1 ou posterior](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="6883d-116">To use these examples, your add-in project must [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="6883d-117">Detecte a exibição ativa da apresentação e manipule o evento ActiveViewChanged</span><span class="sxs-lookup"><span data-stu-id="6883d-117">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="6883d-118">Se você estiver criando um suplemento de conteúdo, será necessário obter o modo de exibição ativo da apresentação e manipular o `ActiveViewChanged` evento, como parte do seu `Office.Initialize` manipulador.</span><span class="sxs-lookup"><span data-stu-id="6883d-118">If you are building a content add-in, you will need to get the presentation's active view and handle the `ActiveViewChanged` event, as part of your `Office.Initialize` handler.</span></span>

> [!NOTE]
> <span data-ttu-id="6883d-119">No PowerPoint Online na Web, o evento [Document.ActiveViewChanged](/javascript/api/office/office.document) nunca será acionado porque o modo de Apresentação de Slides é tratado como uma nova sessão.</span><span class="sxs-lookup"><span data-stu-id="6883d-119">In PowerPoint on the web, the [Document.ActiveViewChanged](/javascript/api/office/office.document) event will never fire as Slide Show mode is treated as a new session.</span></span> <span data-ttu-id="6883d-120">Nesse caso, o suplemento deve obter o modo de exibição ativo ao carregar, conforme observado abaixo.</span><span class="sxs-lookup"><span data-stu-id="6883d-120">In this case, the add-in must fetch the active view on load, as shown in the following code sample.</span></span>

<span data-ttu-id="6883d-121">No seguinte exemplo de código:</span><span class="sxs-lookup"><span data-stu-id="6883d-121">In the following code sample:</span></span>

- <span data-ttu-id="6883d-122">A função `getActiveFileView` chama o método [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) para retornar se o modo de exibição atual da apresentação for "edição" (qualquer um dos modos de exibição em que é possível editar slides, como **Normal** ou **Modo de Exibição de Estrutura de Tópicos**) ou "leitura" ( **Apresentação de Slides** ou **Modo de Exibição de Leitura**).</span><span class="sxs-lookup"><span data-stu-id="6883d-122">The  `getActiveFileView` function calls the [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" (**Slide Show** or **Reading View**).</span></span>

- <span data-ttu-id="6883d-123">A função `registerActiveViewChanged` chama o método [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) para registrar um manipulador para o evento [Document.ActiveViewChanged](/javascript/api/office/office.document).</span><span class="sxs-lookup"><span data-stu-id="6883d-123">The  `registerActiveViewChanged` function calls the [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to register a handler for the [Document.ActiveViewChanged](/javascript/api/office/office.document) event.</span></span>


```js
//general Office.initialize function. Fires on load of the add-in.
Office.initialize = function(){

    //Gets whether the current view is edit or read.
    var currentView = getActiveFileView();

    //register for the active view changed handler
    registerActiveViewChanged();

    //render the content based off of the currentView
    //....
}

function getActiveFileView()
{
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });

}

function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler,
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                app.showNotification("Action failed with error: " + asyncResult.error.message);
            }
            else {
                app.showNotification(asyncResult.status);
            }
        });
}
```

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="6883d-124">Navegue até um determinado slide na apresentação</span><span class="sxs-lookup"><span data-stu-id="6883d-124">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="6883d-125">No exemplo de código a seguir, a função `getSelectedRange` chama o método [Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) para obter o objeto JSON retornado por `asyncResult.value`, que contém uma matriz denominada `slides`.</span><span class="sxs-lookup"><span data-stu-id="6883d-125">In the following code sample, the `getSelectedRange` function calls the [Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method to get the JSON object returned by `asyncResult.value`, which contains an array named `slides`.</span></span> <span data-ttu-id="6883d-126">A matriz `slides` contém índices, ids e títulos do intervalo selecionado slides (ou do slide atual, se vários slides não forem selecionados).</span><span class="sxs-lookup"><span data-stu-id="6883d-126">The `slides` array contains the ids, titles, and indexes of selected range of slides (or of the current slide, if multiple slides are not selected).</span></span> <span data-ttu-id="6883d-127">Ela também salva a id do primeiro slide no intervalo selecionado em uma variável global.</span><span class="sxs-lookup"><span data-stu-id="6883d-127">It also saves the id of the first slide in the selected range to a global variable.</span></span>

```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

<span data-ttu-id="6883d-128">No seguinte exemplo de código, o método `goToFirstSlide` função chamadas a [Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) para navegar até o primeiro slide identificado pela `getSelectedRange` função mostrada anteriormente.</span><span class="sxs-lookup"><span data-stu-id="6883d-128">In the following code sample, the `goToFirstSlide` function calls the [Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) method to navigate to the first slide that was identified by the `getSelectedRange` function shown previously.</span></span>

```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="6883d-129">Navegue entre os slides na apresentação</span><span class="sxs-lookup"><span data-stu-id="6883d-129">Navigate between slides in the presentation</span></span>

<span data-ttu-id="6883d-130">No exemplo de código a seguir, a função `goToSlideByIndex` chama o método `Document.goToByIdAsync` para navegar para o próximo slide da apresentação.</span><span class="sxs-lookup"><span data-stu-id="6883d-130">In the following code sample, the `goToSlideByIndex` function calls the `Document.goToByIdAsync` method to navigate to the next slide in the presentation.</span></span>

```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="6883d-131">Obtenha a URL da apresentação</span><span class="sxs-lookup"><span data-stu-id="6883d-131">Get the URL of the presentation</span></span>

<span data-ttu-id="6883d-132">No seguinte exemplo de código, o método `getFileUrl` função chamadas a [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) para obter a URL do arquivo da apresentação.</span><span class="sxs-lookup"><span data-stu-id="6883d-132">In the following code sample, the  `getFileUrl` function calls the [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) method to get the URL of the presentation file.</span></span>

```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```

## <a name="create-a-presentation"></a><span data-ttu-id="6883d-133">Criar uma apresentação</span><span class="sxs-lookup"><span data-stu-id="6883d-133">Create a presentation</span></span>

<span data-ttu-id="6883d-134">O suplemento pode criar uma nova apresentação separada da instância do PowerPoint, na qual o suplemento está sendo executado atualmente.</span><span class="sxs-lookup"><span data-stu-id="6883d-134">Your add-in can create a new presentation, separate from the PowerPoint instance in which the add-in is currently running.</span></span> <span data-ttu-id="6883d-135">O namespace do PowerPoint tem o método `createPresentation` para essa finalidade.</span><span class="sxs-lookup"><span data-stu-id="6883d-135">The PowerPoint namespace has the `createPresentation` method for this purpose.</span></span> <span data-ttu-id="6883d-136">Quando esse método é chamado, a nova apresentação é aberta imediatamente e exibida em uma nova instância do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="6883d-136">When this method is called, the new presentation is immediately opened and displayed in a new instance of PowerPoint.</span></span> <span data-ttu-id="6883d-137">O suplemento permanece aberto e em execução com a apresentação anterior.</span><span class="sxs-lookup"><span data-stu-id="6883d-137">Your add-in remains open and running with the previous presentation.</span></span>

```js
PowerPoint.createPresentation();
```

<span data-ttu-id="6883d-138">O método `createPresentation` também cria uma cópia de uma apresentação existente.</span><span class="sxs-lookup"><span data-stu-id="6883d-138">The `createPresentation` method can also create a copy of an existing presentation.</span></span> <span data-ttu-id="6883d-139">O método aceita uma representação de cadeia de caracteres codificada em Base64 de um arquivo .pptx como parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="6883d-139">The method accepts a base64-encoded string representation of an .pptx file as an optional parameter.</span></span> <span data-ttu-id="6883d-140">A apresentação resultante será uma cópia desse arquivo, supondo que o argumento da cadeia de caracteres seja um arquivo .pptx válido.</span><span class="sxs-lookup"><span data-stu-id="6883d-140">The resulting presentation will be a copy of that file, assuming the string argument is a valid .pptx file.</span></span> <span data-ttu-id="6883d-141">A classe [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) pode ser usada para converter um arquivo em uma cadeia de caracteres codificada com Base64, como demonstrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6883d-141">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = function (event) {
    // strip off the metadata before the base64-encoded string
    var startIndex = reader.result.toString().indexOf("base64,");
    var copyBase64 = reader.result.toString().substr(startIndex + 7);

    PowerPoint.createPresentation(copyBase64);
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="see-also"></a><span data-ttu-id="6883d-142">Confira também</span><span class="sxs-lookup"><span data-stu-id="6883d-142">See also</span></span>

- [<span data-ttu-id="6883d-143">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="6883d-143">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="6883d-144">Exemplos de Código do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="6883d-144">PowerPoint Code Samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,PowerPoint)
- [<span data-ttu-id="6883d-145">Como salvar o estado e as configurações do suplemento por documento para suplementos de conteúdo e de painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="6883d-145">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="6883d-146">Ler e gravar dados na seleção ativa em um documento ou planilha</span><span class="sxs-lookup"><span data-stu-id="6883d-146">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="6883d-147">Obter todo o documento por meio de um suplemento para PowerPoint ou Word</span><span class="sxs-lookup"><span data-stu-id="6883d-147">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="6883d-148">Usar temas de documentos em seus suplementos do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="6883d-148">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
