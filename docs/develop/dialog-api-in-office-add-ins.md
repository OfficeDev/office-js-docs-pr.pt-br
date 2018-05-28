---
title: Use a API de Caixa de di?logo em seus Suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b026c3c5871372c52d0b44e36c01fc44a3d2bf04
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="use-the-dialog-api-in-your-office-add-ins"></a>Use a API de Caixa de Di?logo em seus Suplementos do Office

Voc? pode usar a [API de Caixa de di?logo](https://dev.office.com/reference/add-ins/shared/officeui) para abrir caixas de di?logo no seu Suplemento do Office. Este artigo fornece orienta??es para usar a API de Caixa de di?logo em seu Suplemento do Office.

> [!NOTE]
> Para informa??es sobre os programas para os quais a API de Caixa de Di?logo tem suporte no momento, confira [Conjuntos de requisitos da API de Caixa de Di?logo](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets). Atualmente, a API de Caixa de Di?logo tem suporte para Word, Excel, PowerPoint e Outlook.

> Um cen?rio prim?rio para as APIs de Caixa de Di?logo ? habilitar a autentica??o com um recurso como o Google ou o Facebook. Se o seu suplemento exigir dados sobre o usu?rio do Office ou seus recursos acess?veis atrav?s do Microsoft Graph, como o Office 365 ou o OneDrive, recomendamos que voc? use a API de logon ?nico sempre que puder. Se voc? usa as APIs para o logon ?nico, ent?o voc? n?o precisar? da API de Caixa de di?logo. Para mais detalhes, consulte [Habilitar o logon ?nico para Suplementos do Office](sso-in-office-add-ins.md).

Considere abrir uma caixa de di?logo em um painel de tarefas ou suplemento de conte?do ou [comando de suplemento](../design/add-in-commands.md) para fazer o seguinte:

- Exibir p?ginas de entrada que n?o podem ser abertas diretamente em um painel de tarefas.
- Fornecer mais espa?o na tela, ou at? uma tela inteira, para algumas tarefas no seu suplemento.
- Hospedar um v?deo que seria muito pequeno se fosse confinado em um painel de tarefas.

> [!NOTE]
> Como a sobreposi??o de elementos de IU n?o s?o recomend?veis, evite abrir uma caixa de di?logo em um painel de tarefas a menos que seu cen?rio o obrigue a fazer isso. Ao considerar como usar a ?rea de superf?cie de um painel de tarefas, observe que pain?is de tarefas podem ter guias. Por exemplo, confira o exemplo [Suplemento do Excel JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).

A imagem abaixo mostra um exemplo de uma caixa de di?logo.

![Comandos de suplemento](../images/auth-o-dialog-open.png)

A caixa de di?logo sempre abre no centro da tela. O usu?rio pode mov?-la e redimension?-la. A janela ? *n?o modal*: o usu?rio pode continuar a interagir com o documento no aplicativo do Office do host e com a p?gina host no painel de tarefas, caso houver uma.

## <a name="dialog-api-scenarios"></a>Cen?rios da API de Caixa de di?logo

As APIs JavaScript para Office t?m suporte para os seguintes cen?rios com um objeto [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) e duas fun??es no [namespace Office.context.ui](https://dev.office.com/reference/add-ins/shared/officeui).

### <a name="open-a-dialog-box"></a>Abrir uma caixa de di?logo.

Para abrir uma caixa de di?logo, seu c?digo no painel de tarefas chama o m?todo [displayDialogAsync](https://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) e transmite a ele a URL do recurso que voc? deseja abrir. Isso geralmente ? uma p?gina, mas pode ser um m?todo controlador em um aplicativo MVC, uma rota, um m?todo de servi?o Web ou qualquer outro recurso. Neste artigo, 'p?gina' ou 'site' refere-se ao recurso na caixa de di?logo. Apresentamos um exemplo de c?digo simples a seguir.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - A URL usa o protocolo HTTP**S**. Isso ? obrigat?rio para todas as p?ginas carregadas em uma caixa di?logo, n?o apenas para a primeira p?gina carregada.
> - O dom?nio ? o mesmo que o dom?nio da p?gina host, que pode ser a p?gina em um painel de tarefas ou o [arquivo de fun??o](https://dev.office.com/reference/add-ins/manifest/functionfile) de um comando de suplemento. Isso ? necess?rio: a p?gina, o m?todo o controlador ou outro recurso que ? passado para o m?todo `displayDialogAsync` deve estar no mesmo dom?nio que a p?gina de host.

Ap?s o carregamento da primeira p?gina (ou de outro recurso), um usu?rio pode ir para qualquer site (ou outro recurso) que usa HTTPS. Tamb?m ? poss?vel criar a primeira p?gina para redirecionar imediatamente para outro site.

Por padr?o, a caixa de di?logo ocupar? 80% da altura e da largura na tela do dispositivo, mas voc? pode definir porcentagens diferentes. Basta transmitir um objeto de configura??o para o m?todo, como mostra o exemplo a seguir.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

Para ver um suplemento de exemplo que faz isso, confira [Exemplo de API de Caixa de di?logo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

Defina os dois valores como 100% para ter uma verdadeira experi?ncia de tela inteira. O m?ximo real ? 99,5%, e a janela ainda poder? ser movida e redimensionada.

> [!NOTE]
> Apenas uma caixa de di?logo pode ser aberta em uma janela do host. Tentar abrir outra caixa de di?logo gera um erro. Portanto, por exemplo, se um usu?rio abrir uma caixa de di?logo no painel de tarefas, ele n?o poder? abrir uma segunda caixa de di?logo em uma p?gina diferente no painel de tarefas. No entanto, quando uma caixa de di?logo ? aberta em um [comando de suplemento](../design/add-in-commands.md), o comando abre um arquivo HTML novo (mas n?o visto) sempre que ele ? selecionado. Isso cria uma nova janela do host (n?o vista) para que cada janela possa iniciar sua pr?pria caixa de di?logo. Para obter mais informa??es, confira [Erros de displayDialogAsync](#errors-from-displaydialogasync).

### <a name="take-advantage-of-a-performance-option-in-office-online"></a>Aproveite uma op??o de desempenho no Office Online

A propriedade `displayInIframe` ? uma propriedade adicional no objeto de configura??o que voc? pode passar para `displayDialogAsync`. Quando essa propriedade for definida como `true` e o suplemento estiver em execu??o em um documento aberto no Office Online, a caixa de di?logo ser? aberta como um iframe flutuante, em vez de uma janela independente, o que faz com que ela seja aberta mais rapidamente. Apresentamos um exemplo a seguir:

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

O valor padr?o ? `false`, que ? o mesmo que omitir a propriedade inteiramente. Se o suplemento n?o estiver sendo executado no Office Online, `displayInIframe` ser? ignorado.

> [!NOTE]
> Voc? **n?o** dever? usar `displayInIframe: true` se a caixa de di?logo redirecionar a qualquer ponto para uma p?gina que n?o possa ser aberta em um iframe. Por exemplo, as p?ginas de entrada de muitos servi?os Web populares, como Google e Conta da Microsoft, n?o podem ser abertas em um iframe.

### <a name="send-information-from-the-dialog-box-to-the-host-page"></a>Envie informa??es da caixa de di?logo para a p?gina host

A caixa de di?logo n?o pode se comunicar com a p?gina host no painel de tarefas, a menos que:

- A p?gina atual na caixa de di?logo esteja no mesmo dom?nio da p?gina host.
- A biblioteca JavaScript do Office seja carregada na p?gina. Como qualquer p?gina que usa a biblioteca JavaScript do Office, o script da p?gina deve atribuir um m?todo ? propriedade `Office.initialize`, embora ele possa ser um m?todo vazio. Para mais detalhes, confira [Iniciar o suplemento](understanding-the-javascript-api-for-office.md#initializing-your-add-in).

O c?digo na p?gina de di?logo use a fun??o `messageParent` para enviar uma mensagem de cadeia de caracteres ou um valor booliano para a p?gina host. A cadeia de caracteres pode ser uma palavra, uma frase, um blob XML, um JSON em formato de cadeia de caracteres ou qualquer outra coisa que possa ser serializada em uma cadeia de caracteres. Apresentamos um exemplo a seguir:

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - A fun??o `messageParent` ? uma das *?nicas* duas APIs do Office que pode ser chamada na caixa de di?logo. A outra ? `Office.context.requirements.isSetSupported`. Para saber mais, confira [Especificar hosts do Office e requisitos da API](specify-office-hosts-and-api-requirements.md).
> - A fun??o `messageParent` s? pode ser chamada em uma p?gina com o mesmo dom?nio (incluindo o protocolo e a porta) da p?gina host.

No pr?ximo exemplo, `googleProfile` ? uma vers?o em formato de cadeia de caracteres do perfil do Google do usu?rio.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

A p?gina host deve ser configurada para receber a mensagem. Voc? pode fazer isso adicionando um par?metro de retorno de chamada ? chamada original de `displayDialogAsync`. O retorno de chamada atribui um manipulador ao evento `DialogMessageReceived`. Apresentamos um exemplo a seguir:

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20},
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);
```

> [!NOTE]
> - O Office transmite um objeto [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) para o retorno de chamada. Ele representa o resultado de tentativas de abrir a caixa de di?logo, mas n?o representa o resultado de eventos na caixa di?logo. Para obter mais informa??es sobre essa distin??o, confira a se??o [Manipular erros e eventos](#handle-errors-and-events).
> - A propriedade `value` do `asyncResult` ? definida como um objeto [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) que existe na p?gina host, n?o no contexto da execu??o da caixa de di?logo.
> - O `processMessage` ? a fun??o que manipula o evento. Voc? pode dar a ele o nome que desejar.
> - A vari?vel `dialog` ? declarada em um escopo mais amplo do que o retorno de chamada porque ela tamb?m ? referenciada em `processMessage`.

Veja a seguir um exemplo simples de um manipulador para o evento `DialogMessageReceived`:

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - O Office transmite o objeto `arg` para o manipulador. Sua propriedade `message` ? o booliano ou a cadeia de caracteres enviada pela chamada de `messageParent` na caixa de di?logo. Neste exemplo, ela ? uma representa??o em formato de cadeia de caracteres de um perfil de usu?rio de um servi?o como a Conta da Microsoft ou o Google, portanto est? desserializada como um objeto com `JSON.parse` novamente.
> - A implementa??o de `showUserName` n?o ? mostrada. Ela pode exibir uma mensagem de boas-vindas personalizada no painel de tarefas.

Quando a intera??o do usu?rio com a caixa de di?logo for conclu?da, seu manipulador de mensagem fechar? a caixa de di?logo, conforme mostrado neste exemplo.

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - O objeto `dialog` deve ser o mesmo que ? retornado pela chamada de `displayDialogAsync`.
> - A chamada de `dialog.close` informa ao Office para fechar a caixa de di?logo imediatamente.

Para ver um suplemento de exemplo que usa essas t?cnicas, confira [Exemplo de API de Caixa de di?logo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

Se o suplemento precisa abrir uma p?gina diferente do painel de tarefas depois de receber a mensagem, ? poss?vel usar o m?todo `window.location.replace` (ou `window.location.href`) como a ?ltima linha do manipulador. Apresentamos um exemplo a seguir:

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

Para ver um exemplo de um suplemento que faz isso, confira o exemplo [Inserir gr?ficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

#### <a name="conditional-messaging"></a>Mensagens condicionais
Como voc? pode enviar v?rias chamadas `messageParent` a partir da caixa de di?logo, mas tem apenas um manipulador na p?gina host do evento `DialogMessageReceived`, o manipulador tem que usar a l?gica condicional para distinguir mensagens diferentes. Por exemplo, se a caixa de di?logo solicitar que o usu?rio entre em um provedor de identidade como a Conta da Microsoft ou o Google, ele enviar? o perfil do usu?rio como uma mensagem. Se a autentica??o falhar, a caixa de di?logo enviar? informa??es de erro ? p?gina host, como no exemplo a seguir:

```js
if (loginSuccess) {
    var userProfile = getProfile();
    var messageObject = {messageType: "signinSuccess", profile: userProfile};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    var errorDetails = getError();
    var messageObject = {messageType: "signinFailure", error: errorDetails};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
> - A vari?vel `loginSuccess` poderia ser inicializada por meio da leitura da resposta HTTP no provedor de identidade.
> - A implementa??o das fun??es `getProfile` e `getError` n?o ? exibida. Cada uma delas obt?m dados de um par?metro de consulta ou do corpo da resposta HTTP.
> - S?o enviados objetos an?nimos de diferentes tipos se a entrada for bem-sucedida ou n?o. Ambos t?m uma propriedade `messageType`, mas um tem uma propriedade `profile` e o outro tem uma propriedade `error`.

Para ver exemplos que usam mensagens condicionais, confira
- [Suplemento do Office que usa o servi?o Auth0 para facilitar o login social](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Suplemento do Office que usa o Servi?o do OAuth.io para Simplificar o Acesso a Servi?os Populares Online](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

O c?digo do manipulador na p?gina host usa o valor da propriedade `messageType` para ramificar como no exemplo a seguir. A fun??o `showUserName` ? a mesma do exemplo anterior e a fun??o `showNotification` exibe o erro na interface do usu?rio da p?gina host.

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "signinSuccess") {
        dialog.close();
        showUserName(messageFromDialog.profile.name);
        window.location.replace("/newPage.html");
    } else {
        dialog.close();
        showNotification("Unable to authenticate user: " + messageFromDialog.error);
    }
}
```

### <a name="closing-the-dialog-box"></a>Feche a caixa de di?logo

Voc? pode implementar um bot?o na caixa de di?logo para fech?-la. Para fazer isso, o manipulador de eventos de clique do bot?o deve usar `messageParent` para informar a p?gina host em que o bot?o foi clicado. Apresentamos um exemplo a seguir:

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

O manipulador de p?gina host de `DialogMessageReceived` poderia chamar `dialog.close`, como neste exemplo. (Veja exemplos anteriores que mostram como o objeto dialog ? inicializado.)


```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

Para ver um exemplo que usa essa t?cnica, confira o [padr?o de design da navega??o do di?logo](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation) no reposit?rio [padr?es de design da experi?ncia do usu?rio para Suplementos do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

Mesmo quando voc? n?o tem sua pr?pria interface de usu?rio de di?logo de fechar, um usu?rio final pode fechar a caixa de di?logo escolhendo a op??o **X** no canto superior direito. Essa a??o aciona o evento `DialogEventReceived`. Se seu painel do host precisar saber quando isso acontece, ele dever? declarar um manipulador para esse evento. Confira a se??o [Erros e eventos na janela de di?logo](#errors-and-events-in-the-dialog-window) para ver os detalhes.

## <a name="handle-errors-and-events"></a>Manipular erros e eventos

Seu c?digo deve manipular duas categorias de eventos:

- Erros retornados pela chamada de `displayDialogAsync` porque n?o foi poss?vel criar a caixa de di?logo.
- Erros e outros eventos na janela de di?logo.

### <a name="errors-from-displaydialogasync"></a>Erros de displayDialogAsync

Al?m dos erros gerais de sistema e de plataforma, tr?s erros s?o espec?ficos para chamar `displayDialogAsync`.

|N?mero do c?digo|Significado|
|:-----|:-----|
|12004|O dom?nio que a URL transmitiu para `displayDialogAsync` n?o ? confi?vel. O dom?nio deve ser o mesmo dom?nio que o da p?gina de host (incluindo o protocolo e o n?mero de porta).|
|12005|A URL passada para `displayDialogAsync` usa o protocolo HTTP. HTTPS ? necess?rio. (Em algumas vers?es do Office, a mensagem de erro retornada com 12005 ? a mesma retornada para 12004.)|
|<span id="12007">12007</span>|Uma caixa de di?logo j? est? aberta na janela do host. Uma janela do host, como um painel de tarefas, s? pode ter uma caixa de di?logo aberta por vez.|

Quando `displayDialogAsync` ? chamado, ele sempre transmite um objeto [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) para sua fun??o de retorno de chamada. Se a chamada for bem-sucedida, ou seja, a janela de di?logo for aberta, a propriedade `value` do objeto `AsyncResult` ser? um objeto [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog). Um exemplo disso encontra-se na se??o [Enviar informa??es da caixa de di?logo para a p?gina host](#send-information-from-the-dialog-box-to-the-host-page). Quando a chamada para `displayDialogAsync` falha, a janela n?o ? criada, a propriedade `status` do objeto `AsyncResult` ? definida como "falha" e a propriedade `error` do objeto ? preenchida. Voc? deve ter sempre um retorno de chamada que testa o `status` e responde quando ? um erro. Veja a seguir um exemplo que simplesmente relata a mensagem de erro independentemente do n?mero do c?digo:

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === "failed") {
        showNotification(asynceResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

### <a name="errors-and-events-in-the-dialog-window"></a>Erros e eventos na janela de di?logo

Tr?s erros e eventos, conhecidos por seus n?meros de c?digos, na caixa de di?logo acionar?o um evento `DialogEventReceived` na p?gina host.

|N?mero do c?digo|Significado|
|:-----|:-----|
|12002|Uma destas op??es:<br> - N?o existe uma p?gina na URL transmitida para `displayDialogAsync`.<br> - A p?gina transmitida para `displayDialogAsync` foi carregada, mas a caixa de di?logo foi direcionada para uma p?gina que ela n?o consegue localizar nem carregar ou foi direcionada para uma URL com sintaxe inv?lida.|
|12003|A caixa de di?logo foi direcionada para uma URL com o protocolo HTTP. HTTPS ? necess?rio.|
|12006|A caixa de di?logo foi fechada, geralmente pelo usu?rio ter escolhido o bot?o **X**.|

Seu c?digo pode atribuir um manipulador para o evento `DialogEventReceived` na chamada para `displayDialogAsync`. Apresentamos um exemplo simples a seguir:

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

Para obter um exemplo de um manipulador para o evento `DialogEventReceived` que cria mensagens de erro personalizadas para cada c?digo de erro, veja o exemplo a seguir:

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

Para ver um suplemento de exemplo que manipula erros dessa forma, confira [Exemplo de API de Caixa de di?logo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).


## <a name="pass-information-to-the-dialog-box"></a>Transmitir informa??es para a caixa di?logo

?s vezes, a p?gina host precisa transmitir informa??es para a caixa de di?logo. Voc? pode fazer isso de duas maneiras principais:

- Adicionar par?metros de consulta ? URL que ? transmitida para `displayDialogAsync`.
- Armazenar as informa??es em outro local que seja acess?vel para a janela do host e para a caixa de di?logo. As duas janelas n?o compartilham um armazenamento de sess?o comum, mas *se elas tiverem o mesmo dom?nio* (incluindo o n?mero da porta, se houver algum), compartilhar?o um [local de armazenamento](http://www.w3schools.com/html/html5_webstorage.asp) comum.

### <a name="use-local-storage"></a>Usar o armazenamento local

Para usar o armazenamento local, seu c?digo chama o m?todo `setItem` do objeto `window.localStorage` na p?gina host antes da chamada `displayDialogAsync`, como no exemplo a seguir:

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

O c?digo na janela de di?logo l? o item quando necess?rio, como no exemplo a seguir:

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

Para ver exemplos de suplementos que usam o armazenamento local dessa forma, confira:

- [Suplemento do Office que usa o servi?o Auth0 para facilitar o login social](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Suplemento do Office que usa o Servi?o do OAuth.io para Simplificar o Acesso a Servi?os Populares Online](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

### <a name="use-query-parameters"></a>Usar par?metros de consulta

O exemplo a seguir mostra como transmitir dados com um par?metro de consulta:

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

Para ver um exemplo que usa essa t?cnica, confira [Inserir gr?ficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

O c?digo na janela de di?logo pode analisar a URL e ler o valor do par?metro.

> [!NOTE]
> O Office adiciona automaticamente um par?metro de consulta chamado `_host_info` ? URL que ? transmitida para `displayDialogAsync`. Ele ? anexado ap?s os par?metros de consulta personalizados, se houver algum. Ele n?o ? anexado ?s URLs subsequentes para as quais a caixa de di?logo navega. No futuro, a Microsoft poder? alterar o conte?do desse valor ou remov?-lo completamente para que seu c?digo n?o consiga l?-lo. O mesmo valor ? adicionado ao armazenamento de sess?o da caixa de di?logo. Novamente, *seu c?digo n?o deve ler nem gravar esse valor*.

## <a name="use-the-dialog-apis-to-show-a-video"></a>Use APIs de Caixa de Di?logo para exibir um v?deo

Para mostrar um v?deo em uma caixa de di?logo:

1.  Crie uma p?gina cujo ?nico conte?do seja um iframe. O atributo `src` dos pontos do iframe para um v?deo online. O protocolo da URL do v?deo deve ser HTTP**S**. Neste artigo, chamaremos esta p?gina de "video.dialogbox.html". Veja a seguir um exemplo da marca??o:

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2.  A p?gina video.dialogbox.html deve estar no mesmo dom?nio que a p?gina de host.
3.  Use uma chamada de `displayDialogAsync` na p?gina host para abrir video.dialogbox.html.
4.  Se o suplemento precisar saber quando o usu?rio fecha a caixa de di?logo, registre um manipulador para o evento `DialogEventReceived` e manipule o evento 12006. Para mais detalhes, confira a se??o [Erros e eventos na janela de di?logo](#errors-and-events-in-the-dialog-window).

Para ver um exemplo que usa um v?deo na caixa de di?logo, confira o [padr?o de design do video placemat](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat) no reposit?rio [padr?es de design da experi?ncia do usu?rio para Suplementos do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

![Captura de tela de um v?deo mostrando uma caixa de di?logo de suplemento](../images/video-placemats-dialog-open.png)

## <a name="use-the-dialog-apis-in-an-authentication-flow"></a>Use as APIs de Caixa de Di?logo em um fluxo de autentica??o

O principal cen?rio das APIs de Caixa de di?logo ? habilitar a autentica??o com um provedor de identidade ou recurso que n?o permite que a p?gina de entrada abra em um Iframe, como uma Conta da Microsoft, o Office 365, o Google e o Facebook.

> [!NOTE]
> Ao usar as APIs de Di?logo para esse cen?rio, *n?o* use a op??o `displayInIframe: true` na chamada para `displayDialogAsync`. Confira [Tirar proveito de uma op??o de desempenho no Office Online](#take-advantage-of-a-performance-option-in-office-online) para obter detalhes sobre essa op??o anteriormente neste artigo.

O que vem a seguir ? um fluxo de autentica??o simples e t?pico:

1. A primeira p?gina que ? aberta na caixa de di?logo ? uma p?gina local (ou outro recurso) que est? hospedada no dom?nio do suplemento; ou seja, o dom?nio da janela do host. Essa p?gina pode ter uma ?nica interface de usu?rio que informa "Aguarde. Estamos redirecionando voc? para a p?gina onde poder? entrar no *NOME-DO-PROVEDOR*." O c?digo nessa p?gina constr?i a URL da p?gina de entrada do provedor de identidade usando as informa??es que s?o transmitidas para a caixa de di?logo, conforme descrito em [Transmitir informa??es para a caixa de di?logo](#pass-information-to-the-dialog-box).
2. A janela de di?logo redireciona ent?o para a p?gina de entrada. A URL inclui um par?metro de consulta que informa o provedor de identidade para redirecionar a janela de di?logo depois que o usu?rio entrar em uma p?gina espec?fica. Neste artigo, chamaremos essa p?gina de "redirectPage.html". (*Essa p?gina deve estar no mesmo dom?nio que a janela do host*, j? que a ?nica maneira de a janela de di?logo transmitir os resultados da tentativa de entrada ? usar uma chamada de `messageParent`, que s? pode ser chamada em uma p?gina com o mesmo dom?nio da janela do host.)
2. O servi?o do provedor de identidade processa a solicita??o GET recebida na janela de di?logo. Se o usu?rio j? estiver conectado, ele imediatamente redirecionar? a janela para redirectPage.html e incluir? os dados do usu?rio como um par?metro de consulta. Se o usu?rio ainda n?o tiver entrado, a p?gina de entrada do provedor aparecer? na janela para que o usu?rio possa entrar. Para a maioria dos provedores, se o usu?rio n?o consegue entrar com ?xito, o provedor mostra uma p?gina de erro na janela de di?logo e n?o redireciona para redirectPage.html. O usu?rio precisa fechar a janela selecionando o **X** no canto. Se o usu?rio entrar com ?xito, a janela de di?logo ser? redirecionada para redirectPage.html e os dados do usu?rio ser?o inclu?dos como um par?metro de consulta.
3. Quando a p?gina redirectPage.html ? aberta, ela chama `messageParent` para relatar o ?xito ou falha na p?gina host e opcionalmente tamb?m informar dados do usu?rio ou dados de erro.
4. O evento `DialogMessageReceived` ? acionado na p?gina host e seu manipulador fecha a janela de di?logo e, opcionalmente, faz outro processamento da mensagem.

Para ver suplementos de exemplo que usam esse padr?o, confira:

- [Inserir gr?ficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart): O recurso que ? inicialmente aberto na janela de di?logo ? um m?todo controlador que n?o tem seu pr?prio modo de exibi??o. Ele redireciona para a p?gina de entrada do Office 365.
- [Autentica??o de Cliente do Office 365 de Suplementos do Office para AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth): O recurso que ? inicialmente aberto na janela de di?logo ? uma p?gina.

#### <a name="support-multiple-identity-providers"></a>Prestar suporte a v?rios provedores de identidade

Se seu suplemento oferece ao usu?rio uma variedade de op??es de provedores, como a Conta da Microsoft, o Google ou o Facebook, voc? precisa de uma primeira p?gina local (confira a se??o anterior) que forne?a uma interface de usu?rio para a escolha de um provedor. A escolha do provedor acionar? a constru??o da URL de entrada e seu redirecionamento.

Para ver um exemplo que usa esse padr?o, confira [Suplemento do Office que usa o Servi?o Auth0 para Facilitar o Login Social](https://github.com/OfficeDev/Office-Add-in-Auth0).

#### <a name="authorization-of-the-add-in-to-an-external-resource"></a>Autoriza??o do suplemento para um recurso externo

Na Web moderna, os aplicativos Web s?o entidades de seguran?a, como os usu?rios, e o aplicativo tem sua pr?pria identidade e permiss?es para recursos online, como o Office 365, Google Plus, Facebook ou LinkedIn. O aplicativo ? registrado no provedor de recursos antes da implanta??o. O registro inclui:

- Uma lista das permiss?es que o aplicativo precisa para usar recursos de um usu?rio.
- Uma URL para a qual o servi?o do recurso deve retornar um token de acesso quando o aplicativo acessa o servi?o.  

Quando um usu?rio invoca uma fun??o no aplicativo que acessa os dados do usu?rio no servi?o do recurso, ele ? solicitado a entrar no servi?o e a conceder ao aplicativo as permiss?es necess?rias para os recursos do usu?rio. Em seguida, o servi?o redireciona a janela de entrada para a URL previamente registrada e transmite o token de acesso. O aplicativo usa o token de acesso para acessar os recursos do usu?rio.

Voc? pode usar as APIs de Caixa de Di?logo para gerenciar esse processo usando um fluxo semelhante ?quele descrito para os usu?rios entrarem. As ?nicas diferen?as s?o:

- Se o usu?rio ainda n?o tiver concedido ao aplicativo as permiss?es necess?rias, ele ser? solicitada a faz?-lo na caixa de di?logo ap?s entrar.
- A janela de di?logo envia o token de acesso ? janela do host usando `messageParent` para enviar o token de acesso em formato de cadeia de caracteres ou armazenando o token de acesso em um local onde a janela do host poder? recuper?-lo. O token tem um limite de tempo, mas enquanto durar, a janela do host poder us?-lo para acessar recursos do usu?rio de forma direta, sem outras solicita??es.

Os exemplos a seguir usam as APIs de Caixa de di?logo para essa finalidade:
- [Inserir gr?ficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart): armazena o token de acesso em um banco de dados.
- [Suplemento do Office que usa o Servi?o do OAuth.io para Simplificar o Acesso a Servi?os Populares Online](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

Para mais informa??es sobre a autentica??o e autoriza??o em suplementos, consulte:
- [Autorizar servi?os externos no seu Suplemento do Office](auth-external-add-ins.md)
- [Biblioteca de Auxiliares da API JavaScript para Office](https://github.com/OfficeDev/office-js-helpers)


## <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>Usar a API de Caixa de di?logo para Office com aplicativos de p?gina ?nica e roteamento do lado do cliente

Se seu suplemento usa o roteamento do lado do cliente, como os aplicativos de p?gina ?nica geralmente fazem, voc? tem a op??o de transmitir a URL de um roteamento para o m?todo [displayDialogAsync](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync), em vez da URL de uma p?gina HTML completa e separada.

> [!IMPORTANT]
>A caixa de di?logo est? em uma nova janela com seu pr?prio contexto de execu??o. Se voc? transmitir uma rota, sua p?gina de base e todos os c?digos de inicializa??o e bootstrapping ser?o executados novamente nesse novo contexto e todas as vari?veis ser?o definidas com seus valores iniciais na caixa de di?logo. Portanto, essa t?cnica inicia uma segunda inst?ncia do aplicativo na janela de di?logo. O c?digo que altera as vari?veis na janela de di?logo n?o altera a vers?o do painel tarefas das mesmas vari?veis. De forma semelhante, a janela de di?logo tem seu pr?prio armazenamento de sess?o que n?o pode ser acessado a partir do c?digo no painel de tarefas.
