---
title: Use a API de Caixa de diálogo em seus Suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 148f4b564169e62f6444e87074c45cb8e4ce5c63
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/25/2018
ms.locfileid: "25005054"
---
# <a name="use-the-dialog-api-in-your-office-add-ins"></a>Use a API de Caixa de Diálogo em seus Suplementos do Office

Você pode usar a [API de Caixa de diálogo](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) para abrir caixas de diálogo no seu Suplemento do Office. Este artigo fornece orientações para usar a API de Caixa de diálogo em seu Suplemento do Office.

> [!NOTE]
> Para informações sobre os programas para os quais a API de Caixa de Diálogo tem suporte no momento, confira [Conjuntos de requisitos da API de Caixa de Diálogo](https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets?view=office-js). Atualmente, a API de Caixa de Diálogo tem suporte para Word, Excel, PowerPoint e Outlook.

> Um cenário primário para as APIs de Caixa de Diálogo é habilitar a autenticação com um recurso como o Google ou o Facebook. Se o seu suplemento exigir dados sobre o usuário do Office ou seus recursos acessíveis através do Microsoft Graph, como o Office 365 ou o OneDrive, recomendamos que você use a API de logon único sempre que puder. Se você usa as APIs para o logon único, então você não precisará da API de Caixa de diálogo. Para mais detalhes, consulte [Habilitar o logon único para Suplementos do Office](sso-in-office-add-ins.md).

Considere abrir uma caixa de diálogo em um painel de tarefas ou suplemento de conteúdo ou [comando de suplemento](../design/add-in-commands.md) para fazer o seguinte:

- Exibir páginas de entrada que não podem ser abertas diretamente em um painel de tarefas.
- Fornecer mais espaço na tela, ou até uma tela inteira, para algumas tarefas no seu suplemento.
- Hospedar um vídeo que seria muito pequeno se fosse confinado em um painel de tarefas.

> [!NOTE]
> Como a sobreposição de elementos de IU não são recomendáveis, evite abrir uma caixa de diálogo em um painel de tarefas a menos que seu cenário o obrigue a fazer isso. Ao considerar como usar a área de superfície de um painel de tarefas, observe que painéis de tarefas podem ter guias. Por exemplo, confira o exemplo [Suplemento do Excel JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).

A imagem abaixo mostra um exemplo de uma caixa de diálogo.

![Comandos de suplemento](../images/auth-o-dialog-open.png)

A caixa de diálogo sempre abre no centro da tela. O usuário pode movê-la e redimensioná-la. A janela é *não modal*: o usuário pode continuar a interagir com o documento no aplicativo do Office do host e com a página host no painel de tarefas, caso houver uma.

## <a name="dialog-api-scenarios"></a>Cenários da API de Caixa de diálogo

As APIs JavaScript para Office têm suporte para os seguintes cenários com um objeto [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) e duas funções no [namespace Office.context.ui](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js).

### <a name="open-a-dialog-box"></a>Abrir uma caixa de diálogo.

Para abrir uma caixa de diálogo, seu código no painel de tarefas chama o método [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) e transmite a ele a URL do recurso que você deseja abrir. Isso geralmente é uma página, mas pode ser um método controlador em um aplicativo MVC, uma rota, um método de serviço Web ou qualquer outro recurso. Neste artigo, 'página' ou 'site' refere-se ao recurso na caixa de diálogo. Apresentamos um exemplo de código simples a seguir.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - A URL usa o protocolo HTTP**S**. Isso é obrigatório para todas as páginas carregadas em uma caixa diálogo, não apenas para a primeira página carregada.
> - A caixa de diálogo domínio do recurso é igual ao domínio da página host, que pode ser a página em um painel de tarefas ou o [arquivo de função](https://docs.microsoft.com/javascript/office/manifest/functionfile?view=office-js) de um comando de suplemento. Isso é obrigatório: a página, o método de controlador ou outro recurso passado para o `displayDialogAsync` método deve estar no mesmo domínio que a página host.

> [!IMPORTANT]
> A página host e os recursos da caixa de diálogo devem ter o mesmo domínio completo. Se você tentar passar `displayDialogAsync` um subdomínio do domínio do suplemento, ele não funcionará. O domínio completo, incluindo qualquer subdomínio, deve corresponder.

Após o carregamento da primeira página (ou de outro recurso), um usuário pode ir para qualquer site (ou outro recurso) que usa HTTPS. Também é possível criar a primeira página para redirecionar imediatamente para outro site.

Por padrão, a caixa de diálogo ocupará 80% da altura e da largura na tela do dispositivo, mas você pode definir porcentagens diferentes. Basta transmitir um objeto de configuração para o método, como mostra o exemplo a seguir.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

Para ver um suplemento de exemplo que faz isso, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

Defina os dois valores como 100% para ter uma verdadeira experiência de tela inteira. O máximo real é 99,5%, e a janela ainda poderá ser movida e redimensionada.

> [!NOTE]
> Apenas uma caixa de diálogo pode ser aberta em uma janela do host. Tentar abrir outra caixa de diálogo gera um erro. Portanto, por exemplo, se um usuário abrir uma caixa de diálogo no painel de tarefas, ele não poderá abrir uma segunda caixa de diálogo em uma página diferente no painel de tarefas. No entanto, quando uma caixa de diálogo é aberta em um [comando de suplemento](../design/add-in-commands.md), o comando abre um arquivo HTML novo (mas não visto) sempre que ele é selecionado. Isso cria uma nova janela do host (não vista) para que cada janela possa iniciar sua própria caixa de diálogo. Para obter mais informações, confira [Erros de displayDialogAsync](#errors-from-displaydialogasync).

### <a name="take-advantage-of-a-performance-option-in-office-online"></a>Aproveite uma opção de desempenho no Office Online

A propriedade `displayInIframe` é uma propriedade adicional no objeto de configuração que você pode passar para `displayDialogAsync`. Quando essa propriedade for definida como `true` e o suplemento estiver em execução em um documento aberto no Office Online, a caixa de diálogo será aberta como um iframe flutuante, em vez de uma janela independente, o que faz com que ela seja aberta mais rapidamente. Apresentamos um exemplo a seguir:

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

O valor padrão é `false`, que é o mesmo que omitir a propriedade inteiramente. Se o suplemento não estiver sendo executado no Office Online, `displayInIframe` será ignorado.

> [!NOTE]
> Você **não** deverá usar `displayInIframe: true` se a caixa de diálogo redirecionar a qualquer ponto para uma página que não possa ser aberta em um iframe. Por exemplo, as páginas de entrada de muitos serviços Web populares, como Google e Conta da Microsoft, não podem ser abertas em um iframe.

### <a name="send-information-from-the-dialog-box-to-the-host-page"></a>Envie informações da caixa de diálogo para a página host

A caixa de diálogo não pode se comunicar com a página host no painel de tarefas, a menos que:

- A página atual na caixa de diálogo esteja no mesmo domínio da página host.
- A biblioteca JavaScript do Office seja carregada na página. Como qualquer página que usa a biblioteca JavaScript do Office, o script da página deve atribuir um método à propriedade `Office.initialize`, embora ele possa ser um método vazio. Para mais detalhes, confira [Iniciar o suplemento](understanding-the-javascript-api-for-office.md#initializing-your-add-in).

O código na página de diálogo use a função `messageParent` para enviar uma mensagem de cadeia de caracteres ou um valor booliano para a página host. A cadeia de caracteres pode ser uma palavra, uma frase, um blob XML, um JSON em formato de cadeia de caracteres ou qualquer outra coisa que possa ser serializada em uma cadeia de caracteres. Apresentamos um exemplo a seguir:

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - A função `messageParent` é uma das *únicas* duas APIs do Office que pode ser chamada na caixa de diálogo. A outra é `Office.context.requirements.isSetSupported`. Para saber mais, confira [Especificar hosts do Office e requisitos da API](specify-office-hosts-and-api-requirements.md).
> - A função `messageParent` só pode ser chamada em uma página com o mesmo domínio (incluindo o protocolo e a porta) da página host.

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
> - O Office transmite um objeto [AsyncResult]() para o retorno de chamada. Ele representa o resultado de tentativas de abrir a caixa de diálogo, mas não representa o resultado de eventos na caixa diálogo. Para obter mais informações sobre essa distinção, confira a seção [Manipular erros e eventos](#handle-errors-and-events).
> - A propriedade `value` do `asyncResult` é definida como um objeto [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) que existe na página host, não no contexto da execução da caixa de diálogo.
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
> - O Office transmite o objeto `arg` para o manipulador. Sua propriedade `message` é o booliano ou a cadeia de caracteres enviada pela chamada de `messageParent` na caixa de diálogo. Neste exemplo, ela é uma representação em formato de cadeia de caracteres de um perfil de usuário de um serviço como a Conta da Microsoft ou o Google, portanto está desserializada como um objeto com `JSON.parse` novamente.
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

Para ver um exemplo de um suplemento que faz isso, confira o exemplo [Inserir gráficos do Excel usando o Microsoft Graph em um suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

#### <a name="conditional-messaging"></a>Mensagens condicionais
Como você pode enviar várias chamadas `messageParent` a partir da caixa de diálogo, mas tem apenas um manipulador na página host do evento `DialogMessageReceived`, o manipulador tem que usar a lógica condicional para distinguir mensagens diferentes. Por exemplo, se a caixa de diálogo solicitar que o usuário entre em um provedor de identidade como a Conta da Microsoft ou o Google, ele enviará o perfil do usuário como uma mensagem. Se a autenticação falhar, a caixa de diálogo enviará informações de erro à página host, como no exemplo a seguir:

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

Para ver exemplos que usam mensagens condicionais, confira
- [Suplemento do Office que usa o serviço Auth0 para facilitar o login social](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Suplemento do Office que usa o Serviço do OAuth.io para Simplificar o Acesso a Serviços Populares Online](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

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

### <a name="closing-the-dialog-box"></a>Feche a caixa de diálogo

Você pode implementar um botão na caixa de diálogo para fechá-la. Para fazer isso, o manipulador de eventos de clique do botão deve usar `messageParent` para informar a página host em que o botão foi clicado. Apresentamos um exemplo a seguir:

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

O manipulador de página host de `DialogMessageReceived` poderia chamar `dialog.close`, como neste exemplo. (Veja exemplos anteriores que mostram como o objeto dialog é inicializado.)


```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

Para ver um exemplo que usa essa técnica, confira o [padrão de design da navegação do diálogo](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation) no repositório [padrões de design da experiência do usuário para Suplementos do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

Mesmo quando você não tem sua própria interface de usuário de diálogo de fechar, um usuário final pode fechar a caixa de diálogo escolhendo a opção **X** no canto superior direito. Essa ação aciona o evento `DialogEventReceived`. Se seu painel do host precisar saber quando isso acontece, ele deverá declarar um manipulador para esse evento. Confira a seção [Erros e eventos na janela de diálogo](#errors-and-events-in-the-dialog-window) para ver os detalhes.

## <a name="handle-errors-and-events"></a>Manipular erros e eventos

Seu código deve manipular duas categorias de eventos:

- Erros retornados pela chamada de `displayDialogAsync` porque não foi possível criar a caixa de diálogo.
- Erros e outros eventos na janela de diálogo.

### <a name="errors-from-displaydialogasync"></a>Erros de displayDialogAsync

Além dos erros gerais de sistema e de plataforma, três erros são específicos para chamar `displayDialogAsync`.

|Número do código|Significado|
|:-----|:-----|
|12004|O domínio que a URL transmitiu para `displayDialogAsync` não é confiável. O domínio deve ser o mesmo domínio que o da página de host (incluindo o protocolo e o número de porta).|
|12005|A URL passada para `displayDialogAsync` usa o protocolo HTTP. HTTPS é necessário. (Em algumas versões do Office, a mensagem de erro retornada com 12005 é a mesma retornada para 12004.)|
|<span id="12007">12007</span>|Uma caixa de diálogo já está aberta na janela do host. Uma janela do host, como um painel de tarefas, só pode ter uma caixa de diálogo aberta por vez.|

Quando `displayDialogAsync` é chamado, ele sempre transmite um objeto [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js) para sua função de retorno de chamada. Se a chamada for bem-sucedida, ou seja, a janela de diálogo for aberta, a propriedade `value` do objeto `AsyncResult` será um objeto [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js). Um exemplo disso encontra-se na seção [Enviar informações da caixa de diálogo para a página host](#send-information-from-the-dialog-box-to-the-host-page). Quando a chamada para `displayDialogAsync` falha, a janela não é criada, a propriedade `status` do objeto `AsyncResult` é definida como "falha" e a propriedade `error` do objeto é preenchida. Você deve ter sempre um retorno de chamada que testa o `status` e responde quando é um erro. Veja a seguir um exemplo que simplesmente relata a mensagem de erro independentemente do número do código:

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

### <a name="errors-and-events-in-the-dialog-window"></a>Erros e eventos na janela de diálogo

Três erros e eventos, conhecidos por seus números de códigos, na caixa de diálogo acionarão um evento `DialogEventReceived` na página host.

|Número do código|Significado|
|:-----|:-----|
|12002|Uma destas opções:<br> - Não existe uma página na URL transmitida para `displayDialogAsync`.<br> - A página transmitida para `displayDialogAsync` foi carregada, mas a caixa de diálogo foi direcionada para uma página que ela não consegue localizar nem carregar ou foi direcionada para uma URL com sintaxe inválida.|
|12003|A caixa de diálogo foi direcionada para uma URL com o protocolo HTTP. HTTPS é necessário.|
|12006|A caixa de diálogo foi fechada, geralmente pelo usuário ter escolhido o botão **X**.|

Seu código pode atribuir um manipulador para o evento `DialogEventReceived` na chamada para `displayDialogAsync`. Apresentamos um exemplo simples a seguir:

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

Para obter um exemplo de um manipulador para o evento `DialogEventReceived` que cria mensagens de erro personalizadas para cada código de erro, veja o exemplo a seguir:

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

Para ver um suplemento de exemplo que manipula erros dessa forma, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).


## <a name="pass-information-to-the-dialog-box"></a>Transmitir informações para a caixa diálogo

Às vezes, a página host precisa transmitir informações para a caixa de diálogo. Você pode fazer isso de duas maneiras principais:

- Adicionar parâmetros de consulta à URL que é transmitida para `displayDialogAsync`.
- Armazenar as informações em outro local que seja acessível para a janela do host e para a caixa de diálogo. As duas janelas não compartilham um armazenamento de sessão comum, mas *se elas tiverem o mesmo domínio* (incluindo o número da porta, se houver algum), compartilharão um [local de armazenamento](https://www.w3schools.com/html/html5_webstorage.asp) comum.

### <a name="use-local-storage"></a>Usar o armazenamento local

Para usar o armazenamento local, seu código chama o método `setItem` do objeto `window.localStorage` na página host antes da chamada `displayDialogAsync`, como no exemplo a seguir:

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

O código na janela de diálogo lê o item quando necessário, como no exemplo a seguir:

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

Para ver exemplos de suplementos que usam o armazenamento local dessa forma, confira:

- [Suplemento do Office que usa o serviço Auth0 para facilitar o login social](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Suplemento do Office que usa o Serviço do OAuth.io para Simplificar o Acesso a Serviços Populares Online](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

### <a name="use-query-parameters"></a>Usar parâmetros de consulta

O exemplo a seguir mostra como transmitir dados com um parâmetro de consulta:

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

Para ver um exemplo que usa essa técnica, confira [Inserir gráficos do Excel usando o Microsoft Graph em um suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

O código na janela de diálogo pode analisar a URL e ler o valor do parâmetro.

> [!NOTE]
> O Office adiciona automaticamente um parâmetro de consulta chamado `_host_info` à URL que é transmitida para `displayDialogAsync`. Ele é anexado após os parâmetros de consulta personalizados, se houver algum. Ele não é anexado às URLs subsequentes para as quais a caixa de diálogo navega. No futuro, a Microsoft poderá alterar o conteúdo desse valor ou removê-lo completamente para que seu código não consiga lê-lo. O mesmo valor é adicionado ao armazenamento de sessão da caixa de diálogo. Novamente, *seu código não deve ler nem gravar esse valor*.

## <a name="use-the-dialog-apis-to-show-a-video"></a>Use APIs de Caixa de Diálogo para exibir um vídeo

Para mostrar um vídeo em uma caixa de diálogo:

1.  Crie uma página cujo único conteúdo seja um iframe. O atributo `src` dos pontos do iframe para um vídeo online. O protocolo da URL do vídeo deve ser HTTP**S**. Neste artigo, chamaremos esta página de "video.dialogbox.html". Veja a seguir um exemplo da marcação:

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2.  A página video.dialogbox.html deve estar no mesmo domínio que a página de host.
3.  Use uma chamada de `displayDialogAsync` na página host para abrir video.dialogbox.html.
4.  Se o suplemento precisar saber quando o usuário fecha a caixa de diálogo, registre um manipulador para o evento `DialogEventReceived` e manipule o evento 12006. Para mais detalhes, confira a seção [Erros e eventos na janela de diálogo](#errors-and-events-in-the-dialog-window).

Para ver um exemplo que usa um vídeo na caixa de diálogo, confira o [padrão de design do video placemat](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat) no repositório [padrões de design da experiência do usuário para Suplementos do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

![Captura de tela de um vídeo mostrando uma caixa de diálogo de suplemento](../images/video-placemats-dialog-open.png)

## <a name="use-the-dialog-apis-in-an-authentication-flow"></a>Use as APIs de Caixa de Diálogo em um fluxo de autenticação

O principal cenário das APIs de Caixa de diálogo é habilitar a autenticação com um provedor de identidade ou recurso que não permite que a página de entrada abra em um Iframe, como uma Conta da Microsoft, o Office 365, o Google e o Facebook.

> [!NOTE]
> Ao usar as APIs de Diálogo para esse cenário, *não* use a opção `displayInIframe: true` na chamada para `displayDialogAsync`. Confira [Tirar proveito de uma opção de desempenho no Office Online](#take-advantage-of-a-performance-option-in-office-online) para obter detalhes sobre essa opção anteriormente neste artigo.

O que vem a seguir é um fluxo de autenticação simples e típico:

1. A primeira página que é aberta na caixa de diálogo é uma página local (ou outro recurso) que está hospedada no domínio do suplemento; ou seja, o domínio da janela do host. Essa página pode ter uma única interface de usuário que informa "Aguarde. Estamos redirecionando você para a página onde poderá entrar no *NOME-DO-PROVEDOR*." O código nessa página constrói a URL da página de entrada do provedor de identidade usando as informações que são transmitidas para a caixa de diálogo, conforme descrito em [Transmitir informações para a caixa de diálogo](#pass-information-to-the-dialog-box).
2. A janela de diálogo redireciona então para a página de entrada. A URL inclui um parâmetro de consulta que informa o provedor de identidade para redirecionar a janela de diálogo depois que o usuário entrar em uma página específica. Neste artigo, chamaremos essa página de "redirectPage.html". (*Essa página deve estar no mesmo domínio que a janela do host*, já que a única maneira de a janela de diálogo transmitir os resultados da tentativa de entrada é usar uma chamada de `messageParent`, que só pode ser chamada em uma página com o mesmo domínio da janela do host.)
2. O serviço do provedor de identidade processa a solicitação GET recebida na janela de diálogo. Se o usuário já estiver conectado, ele imediatamente redirecionará a janela para redirectPage.html e incluirá os dados do usuário como um parâmetro de consulta. Se o usuário ainda não tiver entrado, a página de entrada do provedor aparecerá na janela para que o usuário possa entrar. Para a maioria dos provedores, se o usuário não consegue entrar com êxito, o provedor mostra uma página de erro na janela de diálogo e não redireciona para redirectPage.html. O usuário precisa fechar a janela selecionando o **X** no canto. Se o usuário entrar com êxito, a janela de diálogo será redirecionada para redirectPage.html e os dados do usuário serão incluídos como um parâmetro de consulta.
3. Quando a página redirectPage.html é aberta, ela chama `messageParent` para relatar o êxito ou falha na página host e opcionalmente também informar dados do usuário ou dados de erro.
4. O evento `DialogMessageReceived` é acionado na página host e seu manipulador fecha a janela de diálogo e, opcionalmente, faz outro processamento da mensagem.

Para ver exemplos de suplementos que usam esse padrão, confira:

- [Inserir gráficos do Excel usando o Microsoft Graph em um suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart): o recurso que é aberto inicialmente na janela de diálogo é um método controlador que não tem seu próprio modo de exibição. Ele redireciona para a página de entrada do Office 365.
- [Suplemento do Office de Autenticação de Cliente do Office 365 para AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth): o recurso que é aberto inicialmente na janela de diálogo é uma página.

#### <a name="support-multiple-identity-providers"></a>Prestar suporte a vários provedores de identidade

Se seu suplemento oferece ao usuário uma variedade de opções de provedores, como a Conta da Microsoft, o Google ou o Facebook, você precisa de uma primeira página local (confira a seção anterior) que forneça uma interface de usuário para a escolha de um provedor. A escolha do provedor acionará a construção da URL de entrada e seu redirecionamento.

Para ver um exemplo que usa esse padrão, confira [Suplemento do Office que usa o Serviço Auth0 para Facilitar o Login Social](https://github.com/OfficeDev/Office-Add-in-Auth0).

#### <a name="authorization-of-the-add-in-to-an-external-resource"></a>Autorização do suplemento para um recurso externo

Na Web moderna, os aplicativos Web são entidades de segurança, como os usuários, e o aplicativo tem sua própria identidade e permissões para recursos online, como o Office 365, Google Plus, Facebook ou LinkedIn. O aplicativo é registrado no provedor de recursos antes da implantação. O registro inclui:

- Uma lista das permissões que o aplicativo precisa para usar recursos de um usuário.
- Uma URL para a qual o serviço do recurso deve retornar um token de acesso quando o aplicativo acessa o serviço.  

Quando um usuário invoca uma função no aplicativo que acessa os dados do usuário no serviço do recurso, ele é solicitado a entrar no serviço e a conceder ao aplicativo as permissões necessárias para os recursos do usuário. Em seguida, o serviço redireciona a janela de entrada para a URL previamente registrada e transmite o token de acesso. O aplicativo usa o token de acesso para acessar os recursos do usuário.

Você pode usar as APIs de Caixa de Diálogo para gerenciar esse processo usando um fluxo semelhante àquele descrito para os usuários entrarem. As únicas diferenças são:

- Se o usuário ainda não tiver concedido ao aplicativo as permissões necessárias, ele será solicitada a fazê-lo na caixa de diálogo após entrar.
- A janela de diálogo envia o token de acesso à janela do host usando `messageParent` para enviar o token de acesso em formato de cadeia de caracteres ou armazenando o token de acesso em um local onde a janela do host poderá recuperá-lo. O token tem um limite de tempo, mas enquanto durar, a janela do host poder usá-lo para acessar recursos do usuário de forma direta, sem outras solicitações.

Os exemplos a seguir usam as APIs de Caixa de diálogo para essa finalidade:
- [Inserir gráficos do Excel usando o Microsoft Graph em um suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) - armazena o token de acesso em um banco de dados.
- [Suplemento do Office que usa o Serviço do OAuth.io para Simplificar o Acesso a Serviços Populares Online](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

Para mais informações sobre a autenticação e autorização em suplementos, consulte:
- [Autorizar serviços externos no seu Suplemento do Office](auth-external-add-ins.md)
- [Biblioteca de Auxiliares da API JavaScript para Office](https://github.com/OfficeDev/office-js-helpers)


## <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>Usar a API de Caixa de diálogo para Office com aplicativos de página única e roteamento do lado do cliente

Se seu suplemento usa o roteamento do lado do cliente, como os aplicativos de página única geralmente fazem, você tem a opção de transmitir a URL de um roteamento para o método [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js), em vez da URL de uma página HTML completa e separada.

> [!IMPORTANT]
>A caixa de diálogo está em uma nova janela com seu próprio contexto de execução. Se você transmitir uma rota, sua página de base e todos os códigos de inicialização e bootstrapping serão executados novamente nesse novo contexto e todas as variáveis serão definidas com seus valores iniciais na caixa de diálogo. Portanto, essa técnica inicia uma segunda instância do aplicativo na janela de diálogo. O código que altera as variáveis na janela de diálogo não altera a versão do painel tarefas das mesmas variáveis. De forma semelhante, a janela de diálogo tem seu próprio armazenamento de sessão que não pode ser acessado a partir do código no painel de tarefas.
