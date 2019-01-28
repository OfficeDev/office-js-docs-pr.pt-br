---
title: Suplementos do PowerPoint
description: ''
ms.date: 01/24/2019
localization_priority: Priority
ms.openlocfilehash: da60c87993bc67057aeec6a4e754f57ae376ddd4
ms.sourcegitcommit: b3812245ee1426c299e6484fdd2096a9212ce823
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/25/2019
ms.locfileid: "29539860"
---
# <a name="powerpoint-add-ins"></a>Suplementos do PowerPoint

Você pode usar suplementos do PowerPoint na criação de soluções envolventes para as apresentações de seus usuários em todas as plataformas, incluindo Windows, iOS, Office Online e Mac. Você pode criar dois tipos de comandos de suplementos do PowerPoint:

- Use **suplementos de conteúdo** para adicionar conteúdo dinâmico do HTML5 às suas apresentações. Por exemplo, confira o suplemento [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) que pode ser usado para inserir um diagrama interativo do LucidChart para seu conjunto.

- Use **suplementos do painel de tarefas** para exibir as informações de referência ou inserir dados na apresentação através de um serviço. Por exemplo, confira o suplemento [Imagens Shutterstock](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) que pode ser usado para adicionar fotos profissionais à sua apresentação. 

## <a name="powerpoint-add-in-scenarios"></a>Cenários de suplemento do PowerPoint

Os exemplos de código no artigo mostram algumas tarefas básicas para desenvolver suplementos para o PowerPoint. Além disso, observe o seguinte:

- Para exibir as informações, esses exemplos dependem da função `app.showNotification`, incluída em modelos de projeto de Suplementos do Office do Visual Studio. Se você não estiver usando o Visual Studio para desenvolver seu suplemento, será necessário substituir a função `showNotification` por seu próprio código. 

- Vários desses exemplos também usam um objeto `Globals` que é declarado fora do âmbito destas funções como:   `var Globals = {activeViewHandler:0, firstSlideId:0};`

- Para usar esses exemplos, seu suplemento de projeto deverá [referenciar a biblioteca do Office.js 1.1 ou posterior](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>Detecte a exibição ativa da apresentação e manipule o evento ActiveViewChanged

Se você estiver criando um suplemento de conteúdo, será necessário obter o modo de exibição ativo da apresentação e manipular o `ActiveViewChanged` evento, como parte do seu `Office.Initialize` manipulador.

> [!NOTE]
> No PowerPoint Online, o evento [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document) nunca será acionado porque o modo de Apresentação de Slides é tratado como uma nova sessão. Nesse caso, o suplemento deve obter o modo de exibição ativo ao carregar, conforme observado abaixo.

No seguinte exemplo de código:

- A função `getActiveFileView` chama o método [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document#getactiveviewasync-options--callback-) para retornar se o modo de exibição atual da apresentação for "edição" (qualquer um dos modos de exibição em que é possível editar slides, como **Normal** ou **Modo de Exibição de Estrutura de Tópicos**) ou "leitura" ( **Apresentação de Slides** ou **Modo de Exibição de Leitura**).

- A função `registerActiveViewChanged` chama o método [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) para registrar um manipulador para o evento [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document).


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

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a>Navegue até um determinado slide na apresentação

No seguinte exemplo de código, o método `getSelectedRange` função chamadas do [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) para acessar o objeto JSON retornado pela `asyncResult.value`, que contém uma matriz chamada **slides**. A matriz **slides** contém índices, ids e títulos do intervalo selecionado slides (ou do slide atual, se vários slides não forem selecionados). Ela também salva a id do primeiro slide no intervalo selecionado em uma variável global.

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

No seguinte exemplo de código, o método `goToFirstSlide` função chamadas a [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) para navegar até o primeiro slide identificado pela `getSelectedRange` função mostrada anteriormente.

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

## <a name="navigate-between-slides-in-the-presentation"></a>Navegue entre os slides na apresentação

No seguinte exemplo de código, o método `goToSlideByIndex` função chamadas a **Document.goToByIdAsync** para navegar até o próximo slide na apresentação.

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

## <a name="get-the-url-of-the-presentation"></a>Obtenha a URL da apresentação

No seguinte exemplo de código, o método `getFileUrl` função chamadas a [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) para obter a URL do arquivo da apresentação.

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

## <a name="create-a-presentation"></a>Criar uma apresentação

O suplemento pode criar uma nova apresentação separada da instância do PowerPoint, na qual o suplemento está sendo executado atualmente. O namespace do PowerPoint tem o método `createPresentation` para essa finalidade. Quando esse método é chamado, a nova apresentação é aberta imediatamente e exibida em uma nova instância do PowerPoint. O suplemento permanece aberto e em execução com a apresentação anterior.

```js
PowerPoint.createPresentation();
```

O método `createPresentation` também cria uma cópia de uma apresentação existente. O método aceita uma representação de cadeia de caracteres codificada em Base64 de um arquivo .pptx como parâmetro opcional. A apresentação resultante será uma cópia desse arquivo, supondo que o argumento da cadeia de caracteres seja um arquivo .pptx válido. A classe [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) pode ser usada para converter um arquivo em uma cadeia de caracteres codificada com Base64, como demonstrado no exemplo a seguir.

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = function (event) {
    // strip off the metadata before the base64-encoded string
    var startIndex = event.target.result.indexOf("base64,");
    var copyBase64 = event.target.result.substr(startIndex + 7);

    PowerPoint.createPresentation(copyBase64);
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="see-also"></a>Confira também

- 
  [Exemplos de Código do PowerPoint](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)
- [Como salvar o estado e as configurações do suplemento por documento para suplementos de conteúdo e de painel de tarefas](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [Ler e gravar dados na seleção ativa em um documento ou planilha](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [Obter todo o documento por meio de um suplemento para PowerPoint ou Word](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [Usar temas de documentos em seus suplementos do PowerPoint](use-document-themes-in-your-powerpoint-add-ins.md)