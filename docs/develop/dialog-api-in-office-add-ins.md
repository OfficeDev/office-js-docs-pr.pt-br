---
title: Usar a API da Caixa de Diálogo do Office nos suplementos do Office
description: Saiba as noções básicas sobre como criar uma caixa de diálogo em um Office Add-in.
ms.date: 01/22/2022
ms.localizationpriority: medium
---

# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Usar a API de diálogo do Office em suplementos do Office

Você pode usar a [API de Caixa de diálogo do Office](/javascript/api/office/office.ui) para abrir caixas de diálogo no seu Suplemento do Office. Este artigo fornece orientações para usar a API de Caixa de diálogo em seu Suplemento do Office.

> [!NOTE]
> Para informações sobre os programas para os quais a API de Caixa de Diálogo tem suporte no momento, confira [Conjuntos de requisitos da API de Caixa de Diálogo](../reference/requirement-sets/dialog-api-requirement-sets.md). No momento, a API de Diálogo tem suporte para Excel, PowerPoint e Word. Outlook suporte está incluído em vários conjuntos de requisitos de Caixa&mdash; de Correio para ver a referência da API para obter mais detalhes.

Um cenário fundamental para a API de Caixa de Diálogo é habilitar a autenticação com um recurso como o Google, o Facebook ou o Microsoft Graph. Para saber mais, confira [ autenticação com APIs de Caixa de Diálogo do Office](auth-with-office-dialog-api.md) *depois* que você se familiarizar com este artigo.

Considere abrir uma caixa de diálogo em um painel de tarefas, suplemento de conteúdo ou [comando de suplemento](../design/add-in-commands.md) para fazer o seguinte:

- Exibe páginas de login que não podem ser abertas diretamente em um painel de tarefas.
- Fornecer mais espaço na tela, ou até uma tela inteira, para algumas tarefas no seu suplemento.
- Hospedar um vídeo que seria muito pequeno se fosse confinado em um painel de tarefas.

> [!NOTE]
> Como a sobreposição de elementos de IU não são recomendáveis, evite abrir uma caixa de diálogo em um painel de tarefas a menos que seu cenário o obrigue a fazer isso. Ao considerar como usar a área de superfície de um painel de tarefas, observe que painéis de tarefas podem ter guias. Para ver um exemplo de um painel de tarefas com guias, consulte o [exemplo Excel JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) de complemento.

A imagem abaixo mostra um exemplo de uma caixa de diálogo.

![Captura de tela mostrando a caixa de diálogo com três opções de entrada exibidas em frente ao Word.](../images/auth-o-dialog-open.png)

A caixa de diálogo sempre abre no centro da tela. O usuário pode movê-la e redimensioná-la. A janela é *não-dal* , um usuário pode continuar a interagir com o documento no aplicativo Office e com a página no painel de tarefas, se houver um.

## <a name="open-a-dialog-box-from-a-host-page"></a>Abrir uma caixa de diálogo em uma página de host

As APIs JavaScript para Office incluem um objeto[Dialog](/javascript/api/office/office.dialog) e duas funções no [namespace Office.context.ui](/javascript/api/office/office.ui).

Para abrir uma caixa de diálogo, seu código, geralmente uma página no painel de tarefas chama o método [displayDialogAsync](/javascript/api/office/office.ui) e transmite a ele a URL do recurso que você deseja abrir. A página em que esse método é chamado é conhecida como "página host". Por exemplo, se você chamar esse método no script index.html em um painel de tarefas, index.html será a página do host da caixa de diálogo que o método abre.

O recurso aberto na página de diálogo geralmente é uma página, mas pode ser um método controlador em um aplicativo MVC, uma rota, um método de serviço Web ou qualquer outro recurso. Neste artigo, 'página' ou 'site' refere-se ao recurso na caixa de diálogo. O código a seguir é um exemplo simples.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - A URL usa o protocolo HTTP **S**. Isso é obrigatório para todas as páginas carregadas em uma caixa diálogo, não apenas para a primeira página carregada.
> - A caixa de diálogo é igual ao domínio da página de host, que pode ser a página em um painel de tarefas ou o [arquivo de função](../reference/manifest/functionfile.md) de um comando de suplemento. Isso é necessário: a página, o método do controlador ou outro recurso que é passado para o método `displayDialogAsync` deve estar no mesmo domínio que a página de host.

> [!IMPORTANT]
> A página de host e o recurso que abrem na caixa de diálogo devem ter o mesmo domínio inteiro. Se você tentar passar `displayDialogAsync` para um subdomínio do domínio do suplemento, ele não funcionará. O domínio completo, incluindo qualquer subdomínio, deve corresponder.

Após o carregamento da primeira página (ou de outro recurso), um usuário pode usar links ou outra interface de usuário para qualquer site (ou outro recurso) que usa HTTPS. Também é possível criar a primeira página para redirecionar imediatamente para outro site.

Por padrão, a caixa de diálogo ocupará 80% da altura e da largura na tela do dispositivo, mas você pode definir porcentagens diferentes. Basta transmitir um objeto de configuração para o método, como mostra o exemplo a seguir.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

Para ver um suplemento de exemplo que faz isso, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example). Para obter mais exemplos que usam `displayDialogAsync`, consulte [Samples](#samples).

Defina os dois valores como 100% para ter uma verdadeira experiência de tela inteira. O máximo real é 99,5%, e a janela ainda poderá ser movida e redimensionada.

> [!NOTE]
> Apenas uma caixa de diálogo pode ser aberta em uma janela do host. Tentar abrir outra caixa de diálogo gera um erro. Por exemplo, se um usuário abrir uma caixa de diálogo de um painel de tarefas, ela não poderá abrir uma segunda caixa de diálogo de uma página diferente no painel de tarefas. No entanto, quando uma caixa de diálogo é aberta em um [comando de suplemento](../design/add-in-commands.md), o comando abre um arquivo HTML novo (mas não visto) sempre que ele é selecionado. Isso cria uma nova janela do host (não vista) para que cada janela possa iniciar sua própria caixa de diálogo. Para obter mais informações, confira [Erros de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>Aproveite uma opção de desempenho no Office na Web

A propriedade `displayInIframe` é uma propriedade adicional no objeto de configuração que você pode passar para o`displayDialogAsync`. Quando essa propriedade for definida como `true` e o suplemento estiver em execução em um documento aberto no Office Online, a caixa de diálogo será aberta como um iframe flutuante, em vez de uma janela independente, o que faz com que ela seja aberta mais rapidamente. Apresentamos um exemplo a seguir.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

O valor padrão é `false`, que é o mesmo que omitir a propriedade inteiramente. Se o complemento não estiver sendo executado Office na Web, o `displayInIframe` será ignorado.

> [!NOTE]
> Você não **deve** usar `displayInIframe: true` se a caixa de diálogo redirecionar em qualquer ponto para uma página que não pode ser aberta em um iframe. Por exemplo, as páginas de entrada de muitos serviços Web populares, como a conta do Google e da Microsoft, não podem ser abertas em um iframe.

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>Envie informações da caixa de diálogo para a página host

> [!NOTE]
>
> - Para esclarecer, nesta seção, chamamos a mensagem de [destino da](../reference/manifest/functionfile.md) *página host,* mas estritamente falando, as mensagens estão indo para o tempo de execução *JavaScript* no painel de tarefas (ou o tempo de execução que está hospedando um arquivo de função). A distinção só é significativa no caso de mensagens entre domínios. Para obter mais informações, [mensagens entre domínios para o runtime do host](#cross-domain-messaging-to-the-host-runtime).
> - A caixa de diálogo não pode se comunicar com a página host no painel de tarefas, a menos que Office biblioteca da API JavaScript seja carregada na página. (Como qualquer página que use a biblioteca Office API JavaScript, o script da página deve inicializar o add-in. Para obter detalhes, [consulte Initialize your Office Add-in](initialize-add-in.md).)

O código na caixa de diálogo usa a [função messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) para enviar uma mensagem de cadeia de caracteres para a página host. A cadeia de caracteres pode ser uma palavra, frase, blob XML, JSON stringified ou qualquer outra coisa que possa ser serializada em uma cadeia de caracteres ou lançada em uma cadeia de caracteres. Apresentamos um exemplo a seguir.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true.toString());
}
```

> [!IMPORTANT]
> - A `messageParent` função é uma das *duas* Office APIs JS que podem ser chamadas na caixa de diálogo.
> - A outra API JS que pode ser chamada na caixa de diálogo é `Office.context.requirements.isSetSupported`. Para obter informações sobre ele, consulte [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md). No entanto, na caixa de diálogo, essa API não é suportada Outlook 2016 compra única (ou seja, a versão MSI).

No próximo exemplo, `googleProfile` é uma versão em formato de cadeia de caracteres do perfil do Google do usuário.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

A página host deve ser configurada para receber a mensagem. Você pode fazer isso adicionando um parâmetro de retorno de chamada à chamada original de `displayDialogAsync`. O retorno de chamada atribui um manipulador ao evento `DialogMessageReceived`. Apresentamos um exemplo a seguir.

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
>
> - O Office transmite um objeto [AsyncResult](/javascript/api/office/office.asyncresult) para o retorno de chamada. Ele representa o resultado de tentativas de abrir a caixa de diálogo,  Ela não representa o resultado de eventos na caixa diálogo. Para saber mais sobre essa distinção, confira [Manipular erros e eventos](dialog-handle-errors-events.md).
> - A propriedade `value` do `asyncResult` é definida como um objeto [Dialog](/javascript/api/office/office.dialog) que existe na página host, não no contexto da execução da caixa de diálogo.
> - O `processMessage` é a função que manipula o evento. Você pode dar a ele o nome que desejar.
> - A variável `dialog` é declarada em um escopo mais amplo do que o retorno de chamada porque ela também é referenciada em `processMessage`.

Veja a seguir um exemplo simples de um manipulador para o evento `DialogMessageReceived`.

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
>
> - O Office transmite o objeto `arg` para o manipulador. Sua `message` propriedade é a cadeia de caracteres enviada pela chamada `messageParent` da caixa de diálogo. Neste exemplo, é uma representação stringified do perfil de um usuário de um serviço como a conta da Microsoft ou do Google, portanto, ela é desserializada de volta para um objeto com `JSON.parse`.
> - A `showUserName` implementação não é mostrada. Ela pode exibir uma mensagem de boas-vindas personalizada no painel de tarefas.

Quando a interação do usuário com a caixa de diálogo for concluída, seu manipulador de mensagem fechará a caixa de diálogo, conforme mostrado neste exemplo.

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
>
> - O objeto `dialog` deve ser o mesmo que é retornado pela chamada de `displayDialogAsync`.
> - A chamada de `dialog.close` informa ao Office para fechar a caixa de diálogo imediatamente.

Para ver um suplemento de exemplo que usa essas técnicas, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

Se o suplemento precisa abrir uma página diferente do painel de tarefas depois de receber a mensagem, é possível usar o método `window.location.replace` (ou `window.location.href`) como a última linha do manipulador. Apresentamos um exemplo a seguir.

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

Como você pode enviar várias chamadas `messageParent` a partir da caixa de diálogo, mas tem apenas um manipulador na página host do evento `DialogMessageReceived`, o manipulador tem que usar a lógica condicional para distinguir mensagens diferentes. Por exemplo, se a caixa de diálogo solicitar que um usuário entre em um provedor de identidade, como a conta da Microsoft ou o Google, ela envia o perfil do usuário como uma mensagem. Se a autenticação falhar, a caixa de diálogo enviará informações de erro para a página host, como no exemplo a seguir.

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
>
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
> A `showNotification` implementação não é mostrada no código de exemplo fornecido por este artigo. Um exemplo de como você pode implementar essa função dentro do suplemento, confira [Exemplo do suplemento do Office exemplo do diálogo API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

### <a name="cross-domain-messaging-to-the-host-runtime"></a>Mensagens entre domínios para o tempo de execução do host

A caixa de diálogo ou o tempo de execução javaScript pai (em um painel de tarefas ou em um tempo de execução sem interface do usuário que hospeda um arquivo de função) pode ser navegado para longe do domínio do complemento após a caixa de diálogo ser aberta. Se uma dessas coisas tiver ocorrido, uma chamada falhará, a `messageParent` menos que seu código especifique o domínio do tempo de execução pai. Você faz isso adicionando um [parâmetro DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) à chamada de `messageParent`. Esse objeto tem uma `targetOrigin` propriedade que especifica o domínio para o qual a mensagem deve ser enviada. Se o parâmetro não for usado, Office supõe que o destino seja o mesmo domínio que a caixa de diálogo está hospedando no momento.

> [!NOTE]
> Usar `messageParent` para enviar uma mensagem entre domínios requer o conjunto de requisitos [De origem da caixa de diálogo 1.1](../reference/requirement-sets/dialog-origin-requirement-sets.md). O `DialogMessageOptions` parâmetro é ignorado em versões mais antigas do Office que não suportam o conjunto de requisitos, portanto, o comportamento do método não será afetado se você passá-lo.

A seguir, um exemplo de uso para `messageParent` enviar uma mensagem entre domínios.

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "https://resource.contoso.com" });
```

> [!NOTE]
> O `DialogMessageOptions` parâmetro foi lançado aproximadamente em 19 de julho de 2021. Por cerca de 30 dias após essa data, na Office na Web, `messageParent` `DialogMessageOptions` a primeira vez que for chamada sem o parâmetro e o pai for um domínio diferente da caixa de diálogo, o usuário será solicitado a aprovar o envio de dados para o domínio de destino. Se o usuário aprovar, a resposta do usuário será armazenada em cache por 24 horas. O usuário não é solicitado novamente durante esse período quando `messageParent` é chamado com o mesmo domínio de destino.

Se a mensagem não incluir dados confidenciais, você poderá definir `targetOrigin` como "\*" o que permite que ela seja enviada para qualquer domínio. Apresentamos um exemplo a seguir.

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "*" });
```

> [!TIP]
> O `DialogMessageOptions` parâmetro foi adicionado ao método `messageParent` como um parâmetro necessário em meados de 2021. Os complementos mais antigos que enviam uma mensagem entre domínios com o método não funcionam mais até que sejam atualizados para usar o novo parâmetro. Até que o add-in seja atualizado, somente no Office para Windows, os usuários e administradores do sistema poderão habilitar esses complementos *a* continuar funcionando especificando os domínios confiáveis com uma configuração de registro: **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**. Para fazer isso, crie `.reg` um arquivo com uma extensão, salve-o no computador Windows e clique duas vezes nele para executar. A seguir, um exemplo do conteúdo desse arquivo.
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="pass-information-to-the-dialog-box"></a>Transmitir informações para a caixa diálogo

Seu complemento pode enviar mensagens da página [host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) para uma caixa de diálogo usando [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)).

### <a name="use-messagechild-from-the-host-page"></a>Usar `messageChild()` na página host

Quando você chama a API Office caixa de diálogo para abrir uma caixa de diálogo, um [objeto Dialog](/javascript/api/office/office.dialog) é retornado. Ele deve ser atribuído a uma variável que tenha um escopo maior do que o [método displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) porque o objeto será referenciado por outros métodos. Apresentamos um exemplo a seguir.

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

Este `Dialog` objeto tem um [método messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) que envia qualquer cadeia de caracteres, incluindo dados stringified, para a caixa de diálogo. Isso gera um evento `DialogParentMessageReceived` na caixa de diálogo. Seu código deve manipular esse evento, conforme mostrado na próxima seção.

Considere um cenário no qual a interface do usuário da caixa de diálogo está relacionada à planilha ativa no momento e a posição dessa planilha em relação às outras planilhas. No exemplo a seguir, envia `sheetPropertiesChanged` Excel de planilha para a caixa de diálogo. Nesse caso, a planilha atual é chamada de "Minha Planilha" e é a segunda planilha na pasta de trabalho. Os dados são encapsulados em um objeto e stringified para que possam ser passados para `messageChild`.

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>Manipular DialogParentMessageReceived na caixa de diálogo

No JavaScript da caixa de diálogo, registre um manipulador `DialogParentMessageReceived` para o evento com o método [UI.addHandlerAsync](/javascript/api/office/office.ui#office-office-ui-addhandlerasync-member(1)) . Isso normalmente é feito nos métodos [Office.onReady ou Office.initialize](initialize-add-in.md), conforme mostrado no seguinte. (Um exemplo mais robusto está abaixo.)

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

Em seguida, defina o `onMessageFromParent` manipulador. O código a seguir continua o exemplo da seção anterior. Observe que Office um argumento para o manipulador `message` e que a propriedade do objeto argumento contém a cadeia de caracteres da página host. Neste exemplo, a mensagem é reconvertida para um objeto e jQuery é usada para definir o título superior da caixa de diálogo para corresponder ao novo nome da planilha.

```javascript
function onMessageFromParent(arg) {
    var messageFromParent = JSON.parse(arg.message);
    $('h1').text(messageFromParent.name);
}
```

É uma prática adequada verificar se o manipulador está registrado corretamente. Você pode fazer isso passando um retorno de chamada para o `addHandlerAsync` método. Isso é executado quando a tentativa de registrar o manipulador é concluída. Use o manipulador para registrar ou mostrar um erro se o manipulador não foi registrado com êxito. Apresentamos um exemplo a seguir. Observe que `reportError` é uma função, não definida aqui, que registra ou exibe o erro.

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a>Caixa de diálogo Mensagens condicionais da página pai para a caixa de diálogo

Como você pode fazer `messageChild` várias chamadas da página host, `DialogParentMessageReceived` mas você tem apenas um manipulador na caixa de diálogo do evento, o manipulador deve usar a lógica condicional para distinguir mensagens diferentes. Você pode fazer isso de uma forma que seja precisamente paralela à forma como estruturaria as mensagens condicionais quando a caixa de diálogo está enviando uma mensagem para a página host, conforme descrito em [Conditional messaging](#conditional-messaging).

> [!NOTE]
> Em algumas situações, `messageChild` a API, que faz parte do conjunto de requisitos [DialogApi 1.2](../reference/requirement-sets/dialog-api-requirement-sets.md), pode não ter suporte. Algumas maneiras alternativas para mensagens pai para caixa de diálogo são descritas em Maneiras alternativas de passar mensagens para uma caixa de diálogo de [sua página host](parent-to-dialog.md).

> [!IMPORTANT]
> O [conjunto de requisitos DialogApi 1.2](../reference/requirement-sets/dialog-api-requirement-sets.md) não pode ser especificado na seção **Requisitos** de um manifesto de um complemento. Você terá que verificar se há suporte para DialogApi 1.2 `isSetSupported` em tempo de execução usando o método conforme descrito em Runtime verifica se há suporte ao método e ao conjunto [de requisitos](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). O suporte para requisitos de manifesto está em desenvolvimento.

### <a name="cross-domain-messaging-to-the-dialog-runtime"></a>Mensagens entre domínios para o tempo de execução da caixa de diálogo

A caixa de diálogo ou o tempo de execução javaScript pai (em um painel de tarefas ou em um tempo de execução sem interface do usuário que hospeda um arquivo de função) pode ser navegado para longe do domínio do complemento após a caixa de diálogo ser aberta. Se uma dessas coisas tiver ocorrido, uma chamada `messageChild` falhará, a menos que seu código especifique o domínio do tempo de execução da caixa de diálogo. Você faz isso adicionando um [parâmetro DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) à chamada de `messageChild`. Esse objeto tem uma `targetOrigin` propriedade que especifica o domínio para o qual a mensagem deve ser enviada. Se o parâmetro não for usado, Office supõe que o destino seja o mesmo domínio que o tempo de execução pai está hospedando no momento. 

> [!NOTE]
> Usar `messageChild` para enviar uma mensagem entre domínios requer o conjunto de requisitos [De origem da caixa de diálogo 1.1](../reference/requirement-sets/dialog-origin-requirement-sets.md). O `DialogMessageOptions` parâmetro é ignorado em versões mais antigas do Office que não suportam o conjunto de requisitos, portanto, o comportamento do método não será afetado se você passá-lo.

A seguir, um exemplo de uso para `messageChild` enviar uma mensagem entre domínios.

```js
dialog.messageChild(messageToDialog, { targetOrigin: "https://resource.contoso.com" });
```

Se a mensagem não incluir dados confidenciais, você poderá definir `targetOrigin` como "\*" o que permite que ela seja *enviada* para qualquer domínio. Apresentamos um exemplo a seguir.

```js
dialog.messageChild(messageToDialog, { targetOrigin: "*" });
```

Como o tempo de execução javaScript que está hospedando a caixa de diálogo não pode acessar a seção **AppDomains** do manifesto e, assim, determinar  se o domínio de onde a mensagem vem é confiável, `DialogParentMessageReceived` você deve usar o manipulador para determinar isso. O objeto que é passado para o manipulador contém o domínio atualmente hospedado no pai como sua `origin` propriedade. Veja a seguir um exemplo de como usar a propriedade.

```javascript
function onMessageFromParent(arg) {
    if (arg.origin === "https://addin.fabrikam.com") {
        // process message
    } else {
        dialog.close();
        showNotification("Messages from " + arg.origin + " are not accepted.");
    }
}
```

Por exemplo, seu código pode usar os métodos [Office.onReady ou Office.initialize](initialize-add-in.md) para armazenar uma matriz de domínios confiáveis em uma variável global. Em `arg.origin` seguida, a propriedade pode ser verificada em relação a essa lista no manipulador.

> [!TIP]
> O `DialogMessageOptions` parâmetro foi adicionado ao método `messageChild` como um parâmetro necessário em meados de 2021. Os complementos mais antigos que enviam uma mensagem entre domínios com o método não funcionam mais até que sejam atualizados para usar o novo parâmetro. Até que o add-in seja atualizado, somente no Office para Windows, os usuários e administradores do sistema poderão habilitar esses complementos *a* continuar funcionando especificando os domínios confiáveis com uma configuração de registro: **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**. Para fazer isso, crie `.reg` um arquivo com uma extensão, salve-o no computador Windows e clique duas vezes nele para executar. A seguir, um exemplo do conteúdo desse arquivo.
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="close-the-dialog-box"></a>Fechar a caixa de diálogo

Você pode implementar um botão na caixa de diálogo para fechá-la. Para fazer isso, o manipulador de eventos de clique do botão deve usar `messageParent` para informar a página host em que o botão foi clicado. Apresentamos um exemplo a seguir.

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

### <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>Usar a OFFICE de diálogo com aplicativos de página única e roteamento do lado do cliente

SPAs e o roteamento do lado do cliente devem ser manuseados com cuidado ao usar a API de diálogo do Office. Confira [práticas recomendadas para usar o Office Dialog API em um SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).

### <a name="error-and-event-handling"></a>Manipulação de erros e eventos

Confira [Manipulando erros e eventos na caixa de diálogo do Office](dialog-handle-errors-events.md).

## <a name="next-steps"></a>Próximas etapas

Saiba mais sobre as armadilhas e as práticas recomendadas para a API de diálogo do Office em [práticas recomendadas e regras para a API do Office Dialog](dialog-best-practices.md).

## <a name="samples"></a>Exemplos

Todos os exemplos a seguir usam `displayDialogAsync`. Alguns têm servidores baseados em NodeJS e outros têm servidores baseados em ASP.NET/IIS, mas a lógica de usar o método é a mesma, independentemente de como o lado do servidor do add-in é implementado.

**Noções básicas:**

- [Exemplo da API da caixa de diálogo do suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [Conteúdo de Treinamento / Criação de Complementos (vários exemplos)](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

**Exemplos mais complexos:**

- [Office do Microsoft add-in Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Suplemento do Office Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [SSO do NodeJS do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)
- [Office add-in ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)
- [Office exemplo de monetização saas de complemento](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [Outlook Do Microsoft Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Outlook SSO de complemento](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)
- [Outlook Visualizador de Token de Complemento](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Outlook mensagem acionável do complemento](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [Outlook Compartilhamento de Complementos para OneDrive](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [PowerPoint Add-in Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Excel de tempo de execução compartilhado](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario)
- [Excel QuickBooks do ASPNET de complemento](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Word Add-in JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word Add-in JS SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [OAuth do cliente do AngularJS do Word Add-in](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [Suplemento do Office Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Office de OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [Office código de padrões de design deux de complemento](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
