---
title: Suplementos do PowerPoint
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 390497e74d4dc52b9d400f242850ab72bdb0eabc
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640075"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="b727f-102">Suplementos do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b727f-102">PowerPoint add-ins</span></span>

<span data-ttu-id="b727f-103">Você pode usar os suplementos do PowerPoint para criar soluções envolventes para apresentações dos usuários em todas as plataformas, incluindo Windows, iOS, Office Online e Mac.</span><span class="sxs-lookup"><span data-stu-id="b727f-103">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iOS, Office Online, and Mac. You can create one of two types of add-ins:</span></span> <span data-ttu-id="b727f-104">Você pode criar dois tipos de suplementos do PowerPoint:</span><span class="sxs-lookup"><span data-stu-id="b727f-104">You can create two types of PowerPoint add-ins:</span></span>

- <span data-ttu-id="b727f-p102">Use os **suplementos de conteúdo** para adicionar conteúdo dinâmico do HTML5 às suas apresentações. Por exemplo, confira o suplemento [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) que pode ser usado para inserir um diagrama interativo do LucidChart para seu conjunto.</span><span class="sxs-lookup"><span data-stu-id="b727f-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>

- <span data-ttu-id="b727f-107">Use os **suplementos de painel de tarefas** para exibir as informações de referência ou inserir dados da apresentação através de um serviço.</span><span class="sxs-lookup"><span data-stu-id="b727f-107">Use **task pane add-ins** to bring in reference information or insert data into the slide via a service. For example, see the Shutterstock Images add-in, which you can use to add professional photos to your presentation.</span></span> <span data-ttu-id="b727f-108">Por exemplo, consulte o suplemento [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false), que você pode usar para adicionar fotos profissionais à sua apresentação.</span><span class="sxs-lookup"><span data-stu-id="b727f-108">Use task pane add-ins to bring in reference information or insert data into the slide via a service. For example, see the [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) add-in, which you can use to add professional photos to your presentation.</span></span> 

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="b727f-109">Cenários de suplemento do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b727f-109">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="b727f-110">Os exemplos de código neste artigo demonstram algumas tarefas básicas para o desenvolvimento de suplementos do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b727f-110">The code examples in the article show you some basic tasks for developing content add-ins for PowerPoint.</span></span> <span data-ttu-id="b727f-111">Observe o seguinte:</span><span class="sxs-lookup"><span data-stu-id="b727f-111">Also note the following:</span></span>

- <span data-ttu-id="b727f-112">Para exibir informações, esses exemplos usam a função `app.showNotification`, que está incluída nos modelos de projeto de suplementos do Visual Studio do Office.</span><span class="sxs-lookup"><span data-stu-id="b727f-112">To display information, these examples use the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates.</span></span> <span data-ttu-id="b727f-113">Se você não estiver usando o Visual Studio para desenvolver seu suplemento, você precisa substituirá a função `showNotification` com seu próprio código.</span><span class="sxs-lookup"><span data-stu-id="b727f-113">If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code.</span></span> 

- <span data-ttu-id="b727f-114">Vários desses exemplos também usam um objeto `Globals` declarado além do escopo dessas funções como:   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span><span class="sxs-lookup"><span data-stu-id="b727f-114">Several of these examples also use a `Globals` object that is declared beyond the scope of these functions as:   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

- <span data-ttu-id="b727f-115">Para usar esses exemplos, seu projeto de suplemento deve [fazer referência à biblioteca do Office.js v1.1 ou posterior](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="b727f-115">To use these examples, your add-in project must [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="b727f-116">Detectar a exibição ativa da apresentação e manipular o evento ActiveViewChanged</span><span class="sxs-lookup"><span data-stu-id="b727f-116">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="b727f-117">Se você estiver criando um suplemento de conteúdo, você precisará obter a exibição ativa da apresentação e lidar com o evento `ActiveViewChanged`  como parte do seu manipulador `Office.Initialize`.</span><span class="sxs-lookup"><span data-stu-id="b727f-117">If you are building a content add-in, you will need to get the presentation's active view and handle the ActiveViewChanged event, as part of your Office.Initialize handler.</span></span> 

> [!NOTE]
> <span data-ttu-id="b727f-118">No PowerPoint Online, o evento [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) nunca será acionado como modo de Apresentação de Slides e será tratado como uma nova sessão.</span><span class="sxs-lookup"><span data-stu-id="b727f-118">In PowerPoint Online, the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) event will never fire as Slide Show mode is treated as a new session. In this case, the add-in must fetch the active view on load, as noted below.</span></span> <span data-ttu-id="b727f-119">Nesse caso, o suplemento deve buscar o modo de exibição ativo no carregamento, conforme mostrado no exemplo de código a seguir.</span><span class="sxs-lookup"><span data-stu-id="b727f-119">In this case, the add-in must fetch the active view on load, as shown in the following code sample.</span></span>

<span data-ttu-id="b727f-120">No exemplo de código a seguir:</span><span class="sxs-lookup"><span data-stu-id="b727f-120">In the following code sample:</span></span>

- <span data-ttu-id="b727f-121">A função `getActiveFileView`  chama o método [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-) para retornar se a apresentação atual do modo de exibição é "editar" (qualquer um dos modos de exibição no qual você pode editar os slides, como **Normal** ou **Modo de Estrutura de Tópicos**) ou "ler" (**Apresentação de Slides** ou **Modo de Exibição de Leitura**).</span><span class="sxs-lookup"><span data-stu-id="b727f-121">The  `getActiveFileView` function calls the [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" ( **Slide Show** or **Reading View**) view.</span></span>

- <span data-ttu-id="b727f-122">A função `registerActiveViewChanged` chama o método [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) para registrar um manipulador para o evento [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="b727f-122">The  `registerActiveViewChanged` function calls the [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) method to register a handler for the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) event.</span></span> 


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

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="b727f-123">Navegar até um determinado slide na apresentação</span><span class="sxs-lookup"><span data-stu-id="b727f-123">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="b727f-124">No exemplo de código a seguir, a função `getSelectedRange` chama o método [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) para obter o objeto JSON retornado por `asyncResult.value`, que contém uma matriz denominada **slides**.</span><span class="sxs-lookup"><span data-stu-id="b727f-124">In the following code sample, the `getSelectedRange` function calls the [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) method to get the JSON object returned by `asyncResult.value`, which contains an array named **slides**.</span></span> <span data-ttu-id="b727f-125">A matriz de **slides** contém as ids, os títulos e os índices do intervalo selecionado de slides (ou do slide atual, se vários slides não estiverem selecionados).</span><span class="sxs-lookup"><span data-stu-id="b727f-125">The **slides** array contains the ids, titles, and indexes of selected range of slides (or of the current slide, if multiple slides are not selected).</span></span> <span data-ttu-id="b727f-126">Ela também salva a id do primeiro slide do intervalo selecionado para uma variável global.</span><span class="sxs-lookup"><span data-stu-id="b727f-126">It also saves the id of the first slide in the selected range to a global variable.</span></span>

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

<span data-ttu-id="b727f-127">No exemplo de código a seguir, a função `goToFirstSlide` chama o método [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) para navegar até o primeiro slide que foi identificado pela função `getSelectedRange` mostrada anteriormente.</span><span class="sxs-lookup"><span data-stu-id="b727f-127">In the following code sample, the `goToFirstSlide` function calls the [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) method to navigate to the first slide that was identified by the `getSelectedRange` function shown previously.</span></span>

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

## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="b727f-128">Navegar entre os slides na apresentação</span><span class="sxs-lookup"><span data-stu-id="b727f-128">Navigate between slides in the presentation</span></span>

<span data-ttu-id="b727f-129">No exemplo de código a seguir, a função `goToSlideByIndex` chama o método **Document.goToByIdAsync** para navegar até o próximo slide na apresentação.</span><span class="sxs-lookup"><span data-stu-id="b727f-129">The  `goToSlideByIndex` function calls the **Document.goToByIdAsync** method to navigate to the next slide in the presentation.</span></span>

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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="b727f-130">Obter a URL da apresentação</span><span class="sxs-lookup"><span data-stu-id="b727f-130">Get the URL of the presentation</span></span>

<span data-ttu-id="b727f-131">No exemplo de código a seguir, a função `getFileUrl` chama o método [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) para obter a URL do arquivo de apresentação.</span><span class="sxs-lookup"><span data-stu-id="b727f-131">The  `getFileUrl` function calls the [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) method to get the URL of the presentation file.</span></span>

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



## <a name="see-also"></a><span data-ttu-id="b727f-132">Confira também</span><span class="sxs-lookup"><span data-stu-id="b727f-132">See also</span></span>
- [<span data-ttu-id="b727f-133">Exemplos de código do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b727f-133">PowerPoint Code Samples</span></span>](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)
- [<span data-ttu-id="b727f-134">Como salvar o estado e as configurações do suplemento por documento para suplementos de painel de conteúdo e de tarefa</span><span class="sxs-lookup"><span data-stu-id="b727f-134">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="b727f-135">Ler e gravar dados na seleção ativa em um documento ou planilha</span><span class="sxs-lookup"><span data-stu-id="b727f-135">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="b727f-136">Obter todo o documento por meio de um suplemento para PowerPoint ou Word</span><span class="sxs-lookup"><span data-stu-id="b727f-136">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="b727f-137">Usar temas de documentos em seus suplementos do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b727f-137">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
    
