---
title: Suplementos do PowerPoint
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e5c605410601d711e28ca04ff6e26387019cbb41
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925315"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="27706-102">Suplementos do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="27706-102">PowerPoint add-ins</span></span>

<span data-ttu-id="27706-p101">Você pode usar suplementos do PowerPoint na criação de soluções envolventes para as apresentações de seus usuários em todas as plataformas, incluindo Windows, iOS, Office Online e Mac. Você pode criar um dos dois tipos de suplementos:</span><span class="sxs-lookup"><span data-stu-id="27706-p101">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iOS, Office Online, and Mac. You can create one of two types of add-ins:</span></span>

- <span data-ttu-id="27706-p102">Use **suplementos de conteúdo** para adicionar conteúdo dinâmico do HTML5 às suas apresentações. Por exemplo, confira o suplemento [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) que pode ser usado para inserir um diagrama interativo do LucidChart para seu conjunto.</span><span class="sxs-lookup"><span data-stu-id="27706-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>
- <span data-ttu-id="27706-p103">Use **suplementos do painel de tarefas** para exibir as informações de referência ou inserir dados no slide através de um serviço. Por exemplo, confira o suplemento [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) que pode ser usado para adicionar fotos profissionais à sua apresentação.</span><span class="sxs-lookup"><span data-stu-id="27706-p103">Use **task pane add-ins** to bring in reference information or insert data into the slide via a service. For example, see the [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) add-in, which you can use to add professional photos to your presentation.</span></span> 

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="27706-109">Cenários de suplemento do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="27706-109">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="27706-110">Os exemplos de código no artigo mostram algumas tarefas básicas para desenvolver suplementos de conteúdo para o PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="27706-110">The code examples in the article show you some basic tasks for developing content add-ins for PowerPoint.</span></span> 

<span data-ttu-id="27706-p104">Para exibir as informações, esses exemplos dependem da função `app.showNotification`, incluída em modelos de projeto de Suplementos do Office do Visual Studio. Se você não estiver usando o Visual Studio para desenvolver seu suplemento, será necessário substituir a função `showNotification` por seu próprio código. Vários desses exemplos também dependem desse objeto `globals`, que é declarado fora do escopo destas funções: `var globals = {activeViewHandler:0, firstSlideId:0};`</span><span class="sxs-lookup"><span data-stu-id="27706-p104">To display information, these examples depend on the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates. If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code. Several of these examples also depend on this `globals` object that is declared outside of the scope of these functions: `var globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

<span data-ttu-id="27706-114">Esses exemplos de código exigem que seu projeto faça [referência à biblioteca Office.js v1.1 ou posterior](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="27706-114">These code examples require your project to [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>


## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="27706-115">Detectar a exibição ativa da apresentação e manipular o evento ActiveViewChanged</span><span class="sxs-lookup"><span data-stu-id="27706-115">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="27706-116">Se você estiver criando um suplemento de conteúdo, será necessário obter o modo de exibição ativo da apresentação e manipular o evento ActiveViewChanged como parte do manipulador Office.Initialize.</span><span class="sxs-lookup"><span data-stu-id="27706-116">If you are building a content add-in, you will need to get the presentation's active view and handle the ActiveViewChanged event, as part of your Office.Initialize handler.</span></span>


- <span data-ttu-id="27706-117">A função `getActiveFileView` chama o método [Document.getActiveViewAsync](https://dev.office.com/reference/add-ins/shared/document.getactiveviewasync) para retornar se o modo de exibição atual da apresentação for "edição" (qualquer um dos modos de exibição em que é possível editar slides, como **Normal** ou **Modo de Exibição de Estrutura de Tópicos**) ou "leitura" ( **Apresentação de Slides** ou **Modo de Exibição de Leitura**).</span><span class="sxs-lookup"><span data-stu-id="27706-117">The  `getActiveFileView` function calls the [Document.getActiveViewAsync](https://dev.office.com/reference/add-ins/shared/document.getactiveviewasync) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" ( **Slide Show** or **Reading View**) view.</span></span>


- <span data-ttu-id="27706-118">A função `registerActiveViewChanged` chama o método [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/document.addhandlerasync) para registrar um manipulador para o evento [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged).</span><span class="sxs-lookup"><span data-stu-id="27706-118">The  `registerActiveViewChanged` function calls the [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/document.addhandlerasync) method to register a handler for the [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) event.</span></span> 

> [!NOTE]
> <span data-ttu-id="27706-p105">No PowerPoint Online, o evento [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) nunca será acionado porque o modo de Apresentação de Slides é tratado como uma nova sessão. Nesse caso, o suplemento deve obter o modo de exibição ativo ao carregar, como observado abaixo.</span><span class="sxs-lookup"><span data-stu-id="27706-p105">In PowerPoint Online, the [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) event will never fire as Slide Show mode is treated as a new session. In this case, the add-in must fetch the active view on load, as noted below.</span></span>

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
    

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="27706-121">Navegar até um determinado slide na apresentação</span><span class="sxs-lookup"><span data-stu-id="27706-121">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="27706-p106">A função `getSelectedRange` chama o método [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) para obter um objeto JSON retornado por `asyncResult.value`, que contém uma matriz chamada "slides" contendo as ids, títulos e índices do intervalo selecionado de slides (ou apenas do slide atual). Ela também salva a id do primeiro slide no intervalo selecionado em uma variável global.</span><span class="sxs-lookup"><span data-stu-id="27706-p106">The  `getSelectedRange` function calls the [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) method to get a JSON object returned by `asyncResult.value`, which contains an array named "slides" that contains the ids, titles, and indexes of selected range of slides (or just the current slide). It also saves the id of the first slide in the selected range to a global variable.</span></span>


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

<span data-ttu-id="27706-124">A função `goToFirstSlide` chama o método [Document.goToByIdAsync](https://dev.office.com/reference/add-ins/shared/document.gotobyidasync) para ir até a id do primeiro slide armazenado pela função `getSelectedRange` acima.</span><span class="sxs-lookup"><span data-stu-id="27706-124">The  `goToFirstSlide` function calls the [Document.goToByIdAsync](https://dev.office.com/reference/add-ins/shared/document.gotobyidasync) method to go to the id of the first slide stored by the `getSelectedRange` function above.</span></span>




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


## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="27706-125">Navegar entre os slides na apresentação</span><span class="sxs-lookup"><span data-stu-id="27706-125">Navigate between slides in the presentation</span></span>

<span data-ttu-id="27706-126">A função `goToSlideByIndex` chama o método **Document.goToByIdAsync** para navegar até o próximo slide na apresentação.</span><span class="sxs-lookup"><span data-stu-id="27706-126">The  `goToSlideByIndex` function calls the **Document.goToByIdAsync** method to navigate to the next slide in the presentation.</span></span>


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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="27706-127">Obter a URL da apresentação</span><span class="sxs-lookup"><span data-stu-id="27706-127">Get the URL of the presentation</span></span>

<span data-ttu-id="27706-128">A função `getFileUrl` chama o método [Document.getFileProperties](https://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync) para obter a URL do arquivo da apresentação.</span><span class="sxs-lookup"><span data-stu-id="27706-128">The  `getFileUrl` function calls the [Document.getFileProperties](https://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync) method to get the URL of the presentation file.</span></span>


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



## <a name="see-also"></a><span data-ttu-id="27706-129">Veja também</span><span class="sxs-lookup"><span data-stu-id="27706-129">See also</span></span>
- [<span data-ttu-id="27706-130">Exemplos de Código do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="27706-130">PowerPoint Code Samples</span></span>](https://dev.office.com/code-samples#?filters=powerpoint)
- [<span data-ttu-id="27706-131">Como salvar o estado e as configurações do suplemento por documento para suplementos de conteúdo e de painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="27706-131">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="27706-132">Ler e gravar dados na seleção ativa em um documento ou planilha</span><span class="sxs-lookup"><span data-stu-id="27706-132">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="27706-133">Obter todo o documento por meio de um suplemento para PowerPoint ou Word</span><span class="sxs-lookup"><span data-stu-id="27706-133">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="27706-134">Usar temas de documentos em seus suplementos do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="27706-134">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
    
