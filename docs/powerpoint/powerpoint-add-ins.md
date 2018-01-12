# <a name="powerpoint-add-ins"></a>Suplementos do PowerPoint

Você pode usar suplementos do PowerPoint na criação de soluções envolventes para as apresentações de seus usuários em todas as plataformas, incluindo Windows, iOS, Office Online e Mac. Você pode criar um dos dois tipos de suplementos:

- Use **suplementos de conteúdo** para adicionar conteúdo dinâmico do HTML5 às suas apresentações. Por exemplo, confira o suplemento [LucidChart Diagrams for PowerPoint](https://store.office.com/en-us/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) que pode ser usado para inserir um diagrama interativo do LucidChart para seu conjunto.
- Use **suplementos do painel de tarefas** para exibir as informações de referência ou inserir dados no slide através de um serviço. Por exemplo, confira o suplemento [Shutterstock Images](https://store.office.com/en-us/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) que pode ser usado para adicionar fotos profissionais à sua apresentação. 

>
  **Observação:** Caso pretenda [publicar](../publish/publish.md) o suplemento na Office Store depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação da Office Store](https://msdn.microsoft.com/en-us/library/jj220035.aspx). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) e a [Página de hospedagem e disponibilidade do suplemento do Office](https://dev.office.com/add-in-availability)).

## <a name="powerpoint-add-in-scenarios"></a>Cenários de suplemento do PowerPoint

Os exemplos de código no artigo mostram algumas tarefas básicas para desenvolver suplementos de conteúdo para o PowerPoint. 

Para exibir as informações, esses exemplos dependem da função `app.showNotification`, incluída em modelos de projeto de Suplementos do Office do Visual Studio. Se você não estiver usando o Visual Studio para desenvolver seu suplemento, será necessário substituir a função `showNotification` por seu próprio código. Vários desses exemplos também dependem desse objeto `globals`, que é declarado fora do escopo destas funções: `var globals = {activeViewHandler:0, firstSlideId:0};`

Esses exemplos de código exigem que seu projeto faça [referência à biblioteca Office.js v1.1 ou posterior](../../docs/develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).


## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>Detectar a exibição ativa da apresentação e manipular o evento ActiveViewChanged

Se você estiver criando um suplemento de conteúdo, será necessário obter o modo de exibição ativo da apresentação e manipular o evento ActiveViewChanged como parte do manipulador Office.Initialize.


- A função `getActiveFileView` chama o método [Document.getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md) para retornar se o modo de exibição atual da apresentação for "edição" (qualquer um dos modos de exibição em que é possível editar slides, como **Normal** ou **Modo de Exibição de Estrutura de Tópicos**) ou "leitura" ( **Apresentação de Slides** ou **Modo de Exibição de Leitura**).


- A função `registerActiveViewChanged` chama o método [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) para registrar um manipulador para o evento [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md). 
> Observação: No PowerPoint Online, o evento [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md) nunca será acionado porque o modo de Apresentação de Slides é tratado como uma nova sessão. Nesse caso, o suplemento deve obter o modo de exibição ativo ao carregar, conforme observado abaixo.



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

A função `getSelectedRange` chama o método [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) para obter um objeto JSON retornado por `asyncResult.value`, que contém uma matriz chamada "slides" contendo as ids, títulos e índices do intervalo selecionado de slides (ou apenas do slide atual). Ela também salva a id do primeiro slide no intervalo selecionado em uma variável global.


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

A função `goToFirstSlide` chama o método [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) para ir até a id do primeiro slide armazenado pela função `getSelectedRange` acima.




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

A função `goToSlideByIndex` chama o método **Document.goToByIdAsync** para navegar até o próximo slide na apresentação.


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

A função `getFileUrl` chama o método [Document.getFileProperties](../../reference/shared/document.getfilepropertiesasync.md) para obter a URL do arquivo da apresentação.


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



## <a name="additional-resources"></a>Recursos adicionais
- [Exemplos de Código do PowerPoint](https://dev.office.com/code-samples#?filters=powerpoint)

- [Como salvar o estado e as configurações do suplemento por documento para suplementos de conteúdo e de painel de tarefas](../../docs/develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)

- [Leia e grave dados na seleção ativa, em um documento ou em uma planilha](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [Obter todo o documento por meio de um suplemento para PowerPoint ou Word](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [Usar temas de documentos nos suplementos do PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    
