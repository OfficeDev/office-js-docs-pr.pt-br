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
# <a name="powerpoint-add-ins"></a>Suplementos do PowerPoint

Você pode usar os suplementos do PowerPoint para criar soluções envolventes para apresentações dos usuários em todas as plataformas, incluindo Windows, iOS, Office Online e Mac. Você pode criar dois tipos de suplementos do PowerPoint:

- Use os **suplementos de conteúdo** para adicionar conteúdo dinâmico do HTML5 às suas apresentações. Por exemplo, confira o suplemento [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) que pode ser usado para inserir um diagrama interativo do LucidChart para seu conjunto.

- Use os **suplementos de painel de tarefas** para exibir as informações de referência ou inserir dados da apresentação através de um serviço. Por exemplo, consulte o suplemento [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false), que você pode usar para adicionar fotos profissionais à sua apresentação. 

## <a name="powerpoint-add-in-scenarios"></a>Cenários de suplemento do PowerPoint

Os exemplos de código neste artigo demonstram algumas tarefas básicas para o desenvolvimento de suplementos do PowerPoint. Observe o seguinte:

- Para exibir informações, esses exemplos usam a função `app.showNotification`, que está incluída nos modelos de projeto de suplementos do Visual Studio do Office. Se você não estiver usando o Visual Studio para desenvolver seu suplemento, você precisa substituirá a função `showNotification` com seu próprio código. 

- Vários desses exemplos também usam um objeto `Globals` declarado além do escopo dessas funções como:   `var Globals = {activeViewHandler:0, firstSlideId:0};`

- Para usar esses exemplos, seu projeto de suplemento deve [fazer referência à biblioteca do Office.js v1.1 ou posterior](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>Detectar a exibição ativa da apresentação e manipular o evento ActiveViewChanged

Se você estiver criando um suplemento de conteúdo, você precisará obter a exibição ativa da apresentação e lidar com o evento `ActiveViewChanged`  como parte do seu manipulador `Office.Initialize`. 

> [!NOTE]
> No PowerPoint Online, o evento [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) nunca será acionado como modo de Apresentação de Slides e será tratado como uma nova sessão. Nesse caso, o suplemento deve buscar o modo de exibição ativo no carregamento, conforme mostrado no exemplo de código a seguir.

No exemplo de código a seguir:

- A função `getActiveFileView`  chama o método [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-) para retornar se a apresentação atual do modo de exibição é "editar" (qualquer um dos modos de exibição no qual você pode editar os slides, como **Normal** ou **Modo de Estrutura de Tópicos**) ou "ler" (**Apresentação de Slides** ou **Modo de Exibição de Leitura**).

- A função `registerActiveViewChanged` chama o método [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) para registrar um manipulador para o evento [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js). 


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

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a>Navegar até um determinado slide na apresentação

No exemplo de código a seguir, a função `getSelectedRange` chama o método [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) para obter o objeto JSON retornado por `asyncResult.value`, que contém uma matriz denominada **slides**. A matriz de **slides** contém as ids, os títulos e os índices do intervalo selecionado de slides (ou do slide atual, se vários slides não estiverem selecionados). Ela também salva a id do primeiro slide do intervalo selecionado para uma variável global.

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

No exemplo de código a seguir, a função `goToFirstSlide` chama o método [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) para navegar até o primeiro slide que foi identificado pela função `getSelectedRange` mostrada anteriormente.

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

## <a name="navigate-between-slides-in-the-presentation"></a>Navegar entre os slides na apresentação

No exemplo de código a seguir, a função `goToSlideByIndex` chama o método **Document.goToByIdAsync** para navegar até o próximo slide na apresentação.

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

## <a name="get-the-url-of-the-presentation"></a>Obter a URL da apresentação

No exemplo de código a seguir, a função `getFileUrl` chama o método [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) para obter a URL do arquivo de apresentação.

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



## <a name="see-also"></a>Confira também
- [Exemplos de código do PowerPoint](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)
- [Como salvar o estado e as configurações do suplemento por documento para suplementos de painel de conteúdo e de tarefa](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [Ler e gravar dados na seleção ativa em um documento ou planilha](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [Obter todo o documento por meio de um suplemento para PowerPoint ou Word](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [Usar temas de documentos em seus suplementos do PowerPoint](use-document-themes-in-your-powerpoint-add-ins.md)
    
