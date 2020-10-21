---
title: Usar a API da Caixa de Diálogo do Office nos suplementos do Office
description: Conhecer as noções básicas da criação de uma caixa de diálogo em um suplemento do Office
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 5220d4876d0a8de9c731d2879f0bcb5e669066cd
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626460"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Usar a API de diálogo do Office em suplementos do Office

Você pode usar a [API de Caixa de diálogo do Office](/javascript/api/office/office.ui) para abrir caixas de diálogo no seu Suplemento do Office. Este artigo fornece orientações para usar a API de Caixa de diálogo em seu Suplemento do Office.

> [!NOTE]
> Para informações sobre os programas para os quais a API de Caixa de Diálogo tem suporte no momento, confira [Conjuntos de requisitos da API de Caixa de Diálogo](../reference/requirement-sets/dialog-api-requirement-sets.md). Atualmente, a API de Caixa de Diálogo tem suporte para Word, Excel, PowerPoint e Outlook.

Um cenário fundamental para a API de Caixa de Diálogo é habilitar a autenticação com um recurso como o Google, o Facebook ou o Microsoft Graph. Para saber mais, confira [ autenticação com APIs de Caixa de Diálogo do Office](auth-with-office-dialog-api.md) *depois* que você se familiarizar com este artigo.

Considere abrir uma caixa de diálogo em um painel de tarefas, suplemento de conteúdo ou [comando de suplemento](../design/add-in-commands.md) para fazer o seguinte:

- Exibir páginas de entrada que não podem ser abertas diretamente em um painel de tarefas.
- Fornecer mais espaço na tela, ou até uma tela inteira, para algumas tarefas no seu suplemento.
- Hospedar um vídeo que seria muito pequeno se fosse confinado em um painel de tarefas.

> [!NOTE]
> Como a sobreposição de elementos de IU não são recomendáveis, evite abrir uma caixa de diálogo em um painel de tarefas a menos que seu cenário o obrigue a fazer isso. Ao considerar como usar a área de superfície de um painel de tarefas, observe que painéis de tarefas podem ter guias. Por exemplo, confira o exemplo [Suplemento do Excel JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).

A imagem abaixo mostra um exemplo de uma caixa de diálogo.

![Comandos de suplemento](../images/auth-o-dialog-open.png)

A caixa de diálogo sempre abre no centro da tela. O usuário pode movê-la e redimensioná-la. A janela é não *modal*, e o usuário pode continuar a interagir com o documento no aplicativo do Office e com a página no painel de tarefas, se houver um.

## <a name="open-a-dialog-box-from-a-host-page"></a>Abrir uma caixa de diálogo em uma página de host

As APIs JavaScript para Office incluem um objeto[Dialog](/javascript/api/office/office.dialog) e duas funções no [namespace Office.context.ui](/javascript/api/office/office.ui).

Para abrir uma caixa de diálogo, seu código, geralmente uma página no painel de tarefas chama o método [displayDialogAsync](/javascript/api/office/office.ui) e transmite a ele a URL do recurso que você deseja abrir. A página em que esse método é chamado é conhecida como "página host". Por exemplo, se você chamar esse método no script index.html em um painel de tarefas, index.html será a página do host da caixa de diálogo que o método abre.

O recurso aberto na página de diálogo geralmente é uma página, mas pode ser um método controlador em um aplicativo MVC, uma rota, um método de serviço Web ou qualquer outro recurso. Neste artigo, 'página' ou 'site' refere-se ao recurso na caixa de diálogo. O código a seguir é um exemplo simples:

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - A URL usa o protocolo HTTP**S**. Isso é obrigatório para todas as páginas carregadas em uma caixa diálogo, não apenas para a primeira página carregada.
> - A caixa de diálogo é igual ao domínio da página de host, que pode ser a página em um painel de tarefas ou o [arquivo de função](../reference/manifest/functionfile.md) de um comando de suplemento. Isso é necessário: a página, o método do controlador ou outro recurso que é passado para o método `displayDialogAsync` deve estar no mesmo domínio que a página de host.

> [!IMPORTANT]
> A página de host e o recurso que abrem na caixa de diálogo devem ter o mesmo domínio inteiro. Se você tentar passar `displayDialogAsync` para um subdomínio do domínio do suplemento, ele não funcionará. O domínio completo, incluindo qualquer subdomínio, deve corresponder.

Após o carregamento da primeira página (ou de outro recurso), um usuário pode usar links ou outra interface de usuário para qualquer site (ou outro recurso) que usa HTTPS. Também é possível criar a primeira página para redirecionar imediatamente para outro site.

Por padrão, a caixa de diálogo ocupará 80% da altura e da largura na tela do dispositivo, mas você pode definir porcentagens diferentes. Basta transmitir um objeto de configuração para o método, como mostra o exemplo a seguir.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

Para ver um suplemento de exemplo que faz isso, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

Defina os dois valores como 100% para ter uma verdadeira experiência de tela inteira. O máximo real é 99,5%, e a janela ainda poderá ser movida e redimensionada.

> [!NOTE]
> Apenas uma caixa de diálogo pode ser aberta em uma janela do host. Tentar abrir outra caixa de diálogo gera um erro. Portanto, por exemplo, se um usuário abrir uma caixa de diálogo no painel de tarefas, ele não poderá abrir uma segunda caixa de diálogo em uma página diferente no painel de tarefas. No entanto, quando uma caixa de diálogo é aberta em um [comando de suplemento](../design/add-in-commands.md), o comando abre um arquivo HTML novo (mas não visto) sempre que ele é selecionado. Isso cria uma nova janela do host (não vista) para que cada janela possa iniciar sua própria caixa de diálogo. Para obter mais informações, confira [Erros de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>Aproveite uma opção de desempenho no Office na Web

A propriedade `displayInIframe` é uma propriedade adicional no objeto de configuração que você pode passar para o`displayDialogAsync`. Quando essa propriedade for definida como `true` e o suplemento estiver em execução em um documento aberto no Office Online, a caixa de diálogo será aberta como um iframe flutuante, em vez de uma janela independente, o que faz com que ela seja aberta mais rapidamente. Este é um exemplo:

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

O valor padrão é `false`, que é o mesmo que omitir a propriedade inteiramente. Se o suplemento não estiver sendo executado no Office Online, o `displayInIframe` será ignorado.

> [!NOTE]
> Você **não** deverá usar `displayInIframe: true` se a caixa de diálogo redirecionar a qualquer ponto para uma página que não possa ser aberta em um iframe. Por exemplo, as páginas de entrada de muitos serviços Web populares, como a conta do Google e da Microsoft, não podem ser abertas em um iframe.

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>Envie informações da caixa de diálogo para a página host

A caixa de diálogo não pode se comunicar com a página host no painel de tarefas, a menos que:

- A página atual na caixa de diálogo esteja no mesmo domínio da página host.
- A biblioteca da API JavaScript do Office é carregada na página. (Como qualquer página que usa a biblioteca da API JavaScript do Office, o script para a página deve atribuir um método à `Office.initialize` propriedade, embora possa ser um método vazio. Para obter detalhes, consulte [inicializar o suplemento do Office](initialize-add-in.md).

O código na caixa de diálogo use a função [messageParent](/javascript/api/office/office.ui#messageparent-message-) para enviar uma mensagem de cadeia de caracteres ou um valor booliano para a página host. A cadeia de caracteres pode ser uma palavra, uma frase, um blob XML, um JSON em formato de cadeia de caracteres ou qualquer outra coisa que possa ser serializada em uma cadeia de caracteres. Este é um exemplo:

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!IMPORTANT]
> - A função `messageParent` só pode ser chamada em uma página com o mesmo domínio (incluindo o protocolo e a porta) da página host.
> - A `messageParent` função é uma das *only* duas APIs do Office js que podem ser chamadas na caixa de diálogo. 
> - A outra API JS que pode ser chamada na caixa de diálogo é `Office.context.requirements.isSetSupported` . Para saber mais, confira [especificar requisitos de API e aplicativos do Office](specify-office-hosts-and-api-requirements.md). No entanto, na caixa de diálogo, essa API não tem suporte no Outlook 2016 1-time Purchase (ou seja, a versão MSI).


No próximo exemplo, `googleProfile` é uma versão em formato de cadeia de caracteres do perfil do Google do usuário.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

A página host deve ser configurada para receber a mensagem. Você pode fazer isso adicionando um parâmetro de retorno de chamada à chamada original de `displayDialogAsync`. O retorno de chamada atribui um manipulador ao evento `DialogMessageReceived`. Apresentamos um exemplo a seguir:

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
> - O Office transmite um objeto [AsyncResult](/javascript/api/office/office.asyncresult) para o retorno de chamada. Ele representa o resultado de tentativas de abrir a caixa de diálogo,  Ela não representa o resultado de eventos na caixa diálogo. Para saber mais sobre essa distinção, confira [Manipular erros e eventos](dialog-handle-errors-events.md).
> - A propriedade `value` do `asyncResult` é definida como um objeto [Dialog](/javascript/api/office/office.dialog) que existe na página host, não no contexto da execução da caixa de diálogo.
> - O `processMessage` é a função que manipula o evento. Você pode dar a ele o nome que desejar.
> - A variável `dialog` é declarada em um escopo mais amplo do que o retorno de chamada porque ela também é referenciada em `processMessage`.

Veja a seguir um exemplo simples de um manipulador para o evento `DialogMessageReceived`:

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - O Office transmite o objeto `arg` para o manipulador. Sua propriedade `message` é o booliano ou a cadeia de caracteres enviada pela chamada de `messageParent` na caixa de diálogo. Neste exemplo, é uma representação em formato de um perfil de usuário de um serviço como a conta da Microsoft ou o Google, para que seja desserializado de volta para um objeto com `JSON.parse` .
> - A implementação de `showUserName` não é mostrada. Ela pode exibir uma mensagem de boas-vindas personalizada no painel de tarefas.

Quando a interação do usuário com a caixa de diálogo for concluída, seu manipulador de mensagem fechará a caixa de diálogo, conforme mostrado neste exemplo.

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - O objeto `dialog` deve ser o mesmo que é retornado pela chamada de `displayDialogAsync`.
> - A chamada de `dialog.close` informa ao Office para fechar a caixa de diálogo imediatamente.

Para ver um suplemento de exemplo que usa essas técnicas, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

Se o suplemento precisa abrir uma página diferente do painel de tarefas depois de receber a mensagem, é possível usar o método `window.location.replace` (ou `window.location.href`) como a última linha do manipulador. Apresentamos um exemplo a seguir:

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

Para ver um exemplo de um suplemento que faz isso, consulte [Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

### <a name="conditional-messaging"></a>Mensagens condicionais

Como você pode enviar várias chamadas `messageParent` a partir da caixa de diálogo, mas tem apenas um manipulador na página host do evento `DialogMessageReceived`, o manipulador tem que usar a lógica condicional para distinguir mensagens diferentes. Por exemplo, se a caixa de diálogo solicitar que um usuário entre em um provedor de identidade como a conta da Microsoft ou Google, ele enviará o perfil do usuário como uma mensagem. Se a autenticação falhar, a caixa de diálogo enviará informações de erro à página host, como no exemplo a seguir:

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
> - A variável `loginSuccess` poderia ser inicializada por meio da leitura da resposta HTTP no provedor de identidade.
> - A implementação das funções `getProfile` e `getError` não é exibida. Cada uma delas obtém dados de um parâmetro de consulta ou do corpo da resposta HTTP.
> - São enviados objetos anônimos de diferentes tipos se a entrada for bem-sucedida ou não. Ambos têm uma propriedade `messageType`, mas um tem uma propriedade `profile` e o outro tem uma propriedade `error`.

O código do manipulador na página host usa o valor da propriedade `messageType` para ramificar como no exemplo a seguir. A função `showUserName` é a mesma do exemplo anterior e a função `showNotification` exibe o erro na interface do usuário da página host.

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

> [!NOTE]
> A `showNotification` implementação não é exibida no código de exemplo fornecido neste artigo. Um exemplo de como você pode implementar essa função dentro do suplemento, confira [Exemplo do suplemento do Office exemplo do diálogo API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

## <a name="pass-information-to-the-dialog-box"></a>Transmitir informações para a caixa diálogo

O suplemento pode enviar mensagens da [página de host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) para uma caixa de diálogo usando [Dialog. messageChild](/javascript/api/office/office.dialog#messagechild-message-).

### <a name="use-messagechild-from-the-host-page"></a>Usar `messageChild()` na página host

Quando você chama a API de diálogo do Office para abrir uma caixa de diálogo, um objeto [Dialog](/javascript/api/office/office.dialog) é retornado. Ele deve ser atribuído a uma variável que tenha maior escopo do que o método [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) porque o objeto será referenciado por outros métodos. Este é um exemplo:

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

Este `Dialog` objeto tem um método [messageChild](/javascript/api/office/office.dialog#messagechild-message-) que envia qualquer cadeia de caracteres, incluindo dados em formato, para a caixa de diálogo. Isso gera um `DialogParentMessageReceived` evento na caixa de diálogo. O código deve lidar com esse evento, conforme mostrado na próxima seção.

Considere um cenário em que a interface do usuário da caixa de diálogo está relacionada à planilha ativa no momento e a posição da planilha em relação às outras planilhas. No exemplo a seguir, `sheetPropertiesChanged` envia as propriedades de planilha do Excel para a caixa de diálogo. Nesse caso, a planilha atual é chamada "minha planilha" e é a segunda planilha da pasta de trabalho. Os dados são encapsulados em um objeto e em formato para que possam ser passados `messageChild` .

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>Manipular DialogParentMessageReceived na caixa de diálogo

No JavaScript da caixa de diálogo, registre um manipulador para o `DialogParentMessageReceived` evento com o método [UI. addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) . Isso geralmente é feito nos [métodos Office. onReady ou Office.initialize](initialize-add-in.md), conforme mostrado no seguinte. (Um exemplo mais robusto é o seguinte.)

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

Em seguida, defina o `onMessageFromParent` manipulador. O código a seguir continua o exemplo da seção anterior. Observe que o Office passa um argumento para o manipulador e que a `message` Propriedade do objeto Argument contém a cadeia de caracteres da página host. Neste exemplo, a mensagem é convertida para um objeto e o jQuery é usado para definir o título superior da caixa de diálogo para corresponder ao novo nome da planilha.

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

É uma prática recomendada verificar se o manipulador está registrado corretamente. Você pode fazer isso passando um retorno de chamada para o `addHandlerAsync` método. Isso é executado quando a tentativa de registrar o manipulador é concluída. Use o manipulador para registrar ou mostrar um erro se o manipulador não tiver sido registrado com êxito. Apresentamos um exemplo a seguir. Observe que `reportError` é uma função, não definida aqui, que registra ou exibe o erro.

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a>Mensagem condicional da página pai para a caixa de diálogo

Como você pode fazer várias `messageChild` chamadas a partir da página host, mas tem apenas um manipulador na caixa de diálogo para o `DialogParentMessageReceived` evento, o manipulador deve usar a lógica condicional para distinguir mensagens diferentes. Você pode fazer isso de uma maneira que seja precisamente paralela à forma como você estruturaria mensagens condicionais quando a caixa de diálogo estiver enviando uma mensagem para a página host, conforme descrito em [mensagens condicionais](#conditional-messaging).

> [!NOTE]
> Em algumas situações, a `messageChild` API, que faz parte do conjunto de [requisitos DialogApi 1,2](../reference/requirement-sets/dialog-api-requirement-sets.md), pode não ser suportada. Algumas maneiras alternativas para mensagens de pai para caixa de diálogo são descritas em [maneiras alternativas de passar mensagens para uma caixa de diálogo da página host](parent-to-dialog.md).

> [!IMPORTANT]
> O [conjunto de requisitos DialogApi 1,2](../reference/requirement-sets/dialog-api-requirement-sets.md) não pode ser especificado na `<Requirements>` seção de um manifesto de suplemento. Você precisará verificar o suporte para DialogApi 1,2 em tempo de execução usando o método [isSetSupported](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) . O suporte para requisitos de manifesto está em desenvolvimento.

## <a name="closing-the-dialog-box"></a>Feche a caixa de diálogo

Você pode implementar um botão na caixa de diálogo para fechá-la. Para fazer isso, o manipulador de eventos de clique do botão deve usar `messageParent` para informar a página host em que o botão foi clicado. Apresentamos um exemplo a seguir:

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

O manipulador de página host de `DialogMessageReceived` poderia chamar `dialog.close`, como neste exemplo. (Veja exemplos anteriores que mostram como o objeto `dialog` é inicializado.)

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

Mesmo quando você não tem sua própria interface de usuário de diálogo de fechar, um usuário final pode fechar a caixa de diálogo escolhendo a opção **X** no canto superior direito. Essa ação aciona o evento `DialogEventReceived`. Se seu painel do host precisar saber quando isso acontece, ele deverá declarar um manipulador para esse evento. Confira a seção [Erros e eventos na caixa de diálogo](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box) para ver os detalhes.

## <a name="advanced-topics-and-special-scenarios"></a>Tópicos avançados e cenários especiais

### <a name="use-the-dialog-api-to-show-a-video"></a>Use a API de Caixa de Diálogo para exibir um vídeo

Confira [use a caixa de diálogo do Office para mostrar um vídeo](dialog-video.md).

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a>Use as APIs de Caixa de Diálogo em um fluxo de autenticação

Confira[Autenticar com a API da Caixa de Diálogo do Office](auth-with-office-dialog-api.md).

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>Usar a API de Caixa de diálogo do Office com aplicativos de página única e roteamento do lado do cliente

SPAs e o roteamento do lado do cliente devem ser manuseados com cuidado ao usar a API de diálogo do Office. Confira [práticas recomendadas para usar o Office Dialog API em um SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).

### <a name="error-and-event-handling"></a>Manipulação de erros e eventos

Confira [Manipulando erros e eventos na caixa de diálogo do Office](dialog-handle-errors-events.md).

## <a name="next-steps"></a>Próximas etapas

Saiba mais sobre as armadilhas e as práticas recomendadas para a API de diálogo do Office em [práticas recomendadas e regras para a API do Office Dialog](dialog-best-practices.md).
