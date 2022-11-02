---
title: Usar a API da Caixa de Diálogo do Office nos suplementos do Office
description: Saiba o básico da criação de uma caixa de diálogo em um Suplemento do Office.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4dc1bc0b45bb41952cd2ab83fcd62633d598ab4e
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810012"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Usar a API de diálogo do Office em suplementos do Office

Você pode usar a [API de Caixa de diálogo do Office](/javascript/api/office/office.ui) para abrir caixas de diálogo no seu Suplemento do Office. Este artigo fornece orientações para usar a API de Caixa de diálogo em seu Suplemento do Office.

> [!NOTE]
> Para informações sobre os programas para os quais a API de Caixa de Diálogo tem suporte no momento, confira [Conjuntos de requisitos da API de Caixa de Diálogo](/javascript/api/requirement-sets/common/dialog-api-requirement-sets). Atualmente, há suporte para a API de Caixa de Diálogo para Excel, PowerPoint e Word. O suporte ao Outlook está incluído em vários conjuntos&mdash;de requisitos da Caixa de Correio, consulte a referência de API para obter mais detalhes.

Um cenário fundamental para a API de Caixa de Diálogo é habilitar a autenticação com um recurso como o Google, o Facebook ou o Microsoft Graph. Para saber mais, confira [ autenticação com APIs de Caixa de Diálogo do Office](auth-with-office-dialog-api.md) *depois* que você se familiarizar com este artigo.

Considere abrir uma caixa de diálogo em um painel de tarefas, suplemento de conteúdo ou [comando de suplemento](../design/add-in-commands.md) para fazer o seguinte:

- Exiba páginas de entrada que não podem ser abertas diretamente em um painel de tarefas.
- Fornecer mais espaço na tela, ou até uma tela inteira, para algumas tarefas no seu suplemento.
- Hospedar um vídeo que seria muito pequeno se fosse confinado em um painel de tarefas.

> [!NOTE]
> Como a sobreposição de elementos de IU não são recomendáveis, evite abrir uma caixa de diálogo em um painel de tarefas a menos que seu cenário o obrigue a fazer isso. Ao considerar como usar a área de superfície de um painel de tarefas, observe que painéis de tarefas podem ter guias. Para obter um exemplo de um painel de tarefas com guias, consulte o exemplo [JavaScript SalesTracker do Suplemento do Excel](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) .

A imagem abaixo mostra um exemplo de uma caixa de diálogo.

![Caixa de diálogo com 3 opções de entrada exibidas na frente do Word.](../images/auth-o-dialog-open.png)

A caixa de diálogo sempre abre no centro da tela. O usuário pode movê-la e redimensioná-la. A janela *não é demodal* — um usuário pode continuar a interagir com o documento no aplicativo do Office e com a página no painel de tarefas, se houver um.

## <a name="open-a-dialog-box-from-a-host-page"></a>Abrir uma caixa de diálogo em uma página de host

As APIs JavaScript para Office incluem um objeto[Dialog](/javascript/api/office/office.dialog) e duas funções no [namespace Office.context.ui](/javascript/api/office/office.ui).

Para abrir uma caixa de diálogo, seu código, geralmente uma página no painel de tarefas chama o método [displayDialogAsync](/javascript/api/office/office.ui) e transmite a ele a URL do recurso que você deseja abrir. A página em que esse método é chamado é conhecida como "página host". Por exemplo, se você chamar esse método no script index.html em um painel de tarefas, index.html será a página do host da caixa de diálogo que o método abre.

O recurso aberto na página de diálogo geralmente é uma página, mas pode ser um método controlador em um aplicativo MVC, uma rota, um método de serviço Web ou qualquer outro recurso. Neste artigo, 'página' ou 'site' refere-se ao recurso na caixa de diálogo. O código a seguir é um exemplo simples.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
>
> - A URL usa o protocolo HTTP **S**. Isso é obrigatório para todas as páginas carregadas em uma caixa diálogo, não apenas para a primeira página carregada.
> - A caixa de diálogo é igual ao domínio da página de host, que pode ser a página em um painel de tarefas ou o [arquivo de função](/javascript/api/manifest/functionfile) de um comando de suplemento. Isso é necessário: a página, o método do controlador ou outro recurso que é passado para o método `displayDialogAsync` deve estar no mesmo domínio que a página de host.

> [!IMPORTANT]
> A página de host e o recurso que abrem na caixa de diálogo devem ter o mesmo domínio inteiro. Se você tentar passar `displayDialogAsync` para um subdomínio do domínio do suplemento, ele não funcionará. O domínio completo, incluindo qualquer subdomínio, deve corresponder.

Após o carregamento da primeira página (ou de outro recurso), um usuário pode usar links ou outra interface de usuário para qualquer site (ou outro recurso) que usa HTTPS. Também é possível criar a primeira página para redirecionar imediatamente para outro site.

Por padrão, a caixa de diálogo ocupará 80% da altura e da largura na tela do dispositivo, mas você pode definir porcentagens diferentes. Basta transmitir um objeto de configuração para o método, como mostra o exemplo a seguir.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

Para ver um suplemento de exemplo que faz isso, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example). Para obter mais exemplos que usam `displayDialogAsync`, consulte [Exemplos](#samples).

Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)

> [!NOTE]
> Apenas uma caixa de diálogo pode ser aberta em uma janela do host. Tentar abrir outra caixa de diálogo gera um erro. Por exemplo, se um usuário abrir uma caixa de diálogo de um painel de tarefas, ela não poderá abrir uma segunda caixa de diálogo de uma página diferente no painel de tarefas. No entanto, quando uma caixa de diálogo é aberta em um [comando de suplemento](../design/add-in-commands.md), o comando abre um arquivo HTML novo (mas não visto) sempre que ele é selecionado. Isso cria uma nova janela do host (não vista) para que cada janela possa iniciar sua própria caixa de diálogo. Para obter mais informações, confira [Erros de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>Aproveite uma opção de desempenho no Office na Web

A propriedade `displayInIframe` é uma propriedade adicional no objeto de configuração que você pode passar para o`displayDialogAsync`. Quando essa propriedade for definida como `true` e o suplemento estiver em execução em um documento aberto no Office Online, a caixa de diálogo será aberta como um iframe flutuante, em vez de uma janela independente, o que faz com que ela seja aberta mais rapidamente. Apresentamos um exemplo a seguir.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

O valor padrão é `false`, que é o mesmo que omitir a propriedade inteiramente. Se o suplemento não estiver em execução no Office na Web, o `displayInIframe` será ignorado.

> [!NOTE]
> Você **não** deve usar `displayInIframe: true` se a caixa de diálogo em qualquer ponto redirecionar para uma página que não pode ser aberta em um iframe. Por exemplo, as páginas de entrada de muitos serviços Web populares, como Google e Microsoft, não podem ser abertas em um iframe.

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>Envie informações da caixa de diálogo para a página host

> [!NOTE]
>
> - Para obter clareza, nesta seção chamamos a mensagem de destino da *página* do host, mas estritamente falando as mensagens estão indo para o [Runtime](../testing/runtimes.md) no painel de tarefas (ou o runtime que está hospedando um [arquivo de função](/javascript/api/manifest/functionfile)). A distinção só é significativa no caso de mensagens entre domínios. Para obter mais informações, [mensagens entre domínios para o runtime do host](#cross-domain-messaging-to-the-host-runtime).
> - A caixa de diálogo não pode se comunicar com a página host no painel de tarefas, a menos que a biblioteca de API JavaScript do Office seja carregada na página. (Como qualquer página que usa a biblioteca de API JavaScript do Office, o script para a página deve inicializar o suplemento. Para obter detalhes, consulte [Inicializar seu Suplemento do Office](initialize-add-in.md).)

O código na caixa de diálogo usa a função [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) para enviar uma mensagem de cadeia de caracteres para a página do host. A cadeia de caracteres pode ser uma palavra, frase, blob XML, JSON stringified ou qualquer outra coisa que possa ser serializada para uma cadeia de caracteres ou lançada para uma cadeia de caracteres. Apresentamos um exemplo a seguir.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true.toString());
}
```

> [!IMPORTANT]
>
> - A `messageParent` função é uma das duas *únicas* APIs JS do Office que podem ser chamadas na caixa de diálogo.
> - A outra API JS que pode ser chamada na caixa de diálogo é `Office.context.requirements.isSetSupported`. Para obter informações sobre isso, consulte [Especificar aplicativos do Office e requisitos de API](specify-office-hosts-and-api-requirements.md). No entanto, na caixa de diálogo, essa API não tem suporte no Outlook 2016 perpétuo licenciado por volume (ou seja, na versão MSI).

No próximo exemplo, `googleProfile` é uma versão em formato de cadeia de caracteres do perfil do Google do usuário.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

A página host deve ser configurada para receber a mensagem. Você pode fazer isso adicionando um parâmetro de retorno de chamada à chamada original de `displayDialogAsync`. O retorno de chamada atribui um manipulador ao evento `DialogMessageReceived`. Apresentamos um exemplo a seguir.

```js
let dialog;
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
> - The `processMessage` is the function that handles the event. You can give it any name you want.
> - A variável `dialog` é declarada em um escopo mais amplo do que o retorno de chamada porque ela também é referenciada em `processMessage`.

Veja a seguir um exemplo simples de um manipulador para o evento `DialogMessageReceived`.

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
>
> - O Office transmite o objeto `arg` para o manipulador. Sua `message` propriedade é a cadeia de caracteres enviada pela chamada da caixa de `messageParent` diálogo. Neste exemplo, ele é uma representação em cadeia de caracteres do perfil de um usuário de um serviço como conta Microsoft ou Google, portanto, ele é desserializado de volta para um objeto com `JSON.parse`.
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

If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example.

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

Como você pode enviar várias chamadas `messageParent` a partir da caixa de diálogo, mas tem apenas um manipulador na página host do evento `DialogMessageReceived`, o manipulador tem que usar a lógica condicional para distinguir mensagens diferentes. Por exemplo, se a caixa de diálogo solicitar que um usuário entre em um provedor de identidade, como conta microsoft ou Google, ele enviará o perfil do usuário como uma mensagem. Se a autenticação falhar, a caixa de diálogo enviará informações de erro para a página do host, como no exemplo a seguir.

```js
if (loginSuccess) {
    const userProfile = getProfile();
    const messageObject = {messageType: "signinSuccess", profile: userProfile};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    const errorDetails = getError();
    const messageObject = {messageType: "signinFailure", error: errorDetails};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
>
> - A variável `loginSuccess` poderia ser inicializada por meio da leitura da resposta HTTP no provedor de identidade.
> - A implementação das `getProfile` funções e `getError` não é mostrada. Cada uma delas obtém dados de um parâmetro de consulta ou do corpo da resposta HTTP.
> - Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.

The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
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

### <a name="cross-domain-messaging-to-the-host-runtime"></a>Mensagens entre domínios para o runtime do host

Depois que a caixa de diálogo for aberta, a caixa de diálogo ou o runtime pai poderão navegar para longe do domínio do suplemento. Se alguma dessas coisas acontecer, uma chamada de `messageParent` falhará, a menos que seu código especifique o domínio do runtime pai. Você faz isso adicionando um parâmetro [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) à chamada de `messageParent`. Esse objeto tem uma `targetOrigin` propriedade que especifica o domínio para o qual a mensagem deve ser enviada. Se o parâmetro não for usado, o Office pressupõe que o destino seja o mesmo domínio que a caixa de diálogo está hospedando no momento.

> [!NOTE]
> Usar `messageParent` para enviar uma mensagem entre domínios requer o [conjunto de requisitos Origem do Diálogo 1.1](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets). O `DialogMessageOptions` parâmetro é ignorado em versões mais antigas do Office que não dão suporte ao conjunto de requisitos, portanto, o comportamento do método não será afetado se você passá-lo.

A seguir está um exemplo de uso `messageParent` para enviar uma mensagem entre domínios.

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "https://resource.contoso.com" });
```

> [!NOTE]
> O `DialogMessageOptions` parâmetro foi lançado aproximadamente em 19 de julho de 2021. Por cerca de 30 dias após essa data, em Office na Web, a primeira vez que `messageParent` é chamado sem o `DialogMessageOptions` parâmetro e o pai é um domínio diferente da caixa de diálogo, o usuário será solicitado a aprovar o envio de dados para o domínio de destino. Se o usuário aprovar, a resposta do usuário será armazenada em cache por 24 horas. O usuário não é solicitado novamente durante esse período quando `messageParent` é chamado com o mesmo domínio de destino.

Se a mensagem não incluir dados confidenciais, você poderá definir o `targetOrigin` como "\*" que permite que ele seja enviado para qualquer domínio. Apresentamos um exemplo a seguir.

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "*" });
```

> [!TIP]
> O `DialogMessageOptions` parâmetro foi adicionado ao `messageParent` método como um parâmetro necessário em meados de 2021. Os suplementos mais antigos que enviam uma mensagem entre domínios com o método não funcionam mais até que sejam atualizados para usar o novo parâmetro. Até que o suplemento seja atualizado, *somente no Office no Windows*, usuários e administradores do sistema podem permitir que esses suplementos continuem funcionando especificando os domínios confiáveis com uma configuração de registro: **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**. Para fazer isso, crie um arquivo com uma `.reg` extensão, salve-o no computador Windows e clique duas vezes nele para executá-lo. A seguir está um exemplo do conteúdo de tal arquivo.
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="pass-information-to-the-dialog-box"></a>Transmitir informações para a caixa diálogo

Seu suplemento pode enviar mensagens da página do [host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) para uma caixa de diálogo usando [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)).

### <a name="use-messagechild-from-the-host-page"></a>Usar `messageChild()` na página host

Quando você chama a API da caixa de diálogo do Office para abrir uma caixa de diálogo, um objeto [Dialog](/javascript/api/office/office.dialog) é retornado. Ele deve ser atribuído a uma variável que tenha um escopo maior que o método [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) porque o objeto será referenciado por outros métodos. Apresentamos um exemplo a seguir.

```javascript
let dialog;
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

Esse `Dialog` objeto tem um método [messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) que envia qualquer cadeia de caracteres, incluindo dados stringified, para a caixa de diálogo. Isso gera um `DialogParentMessageReceived` evento na caixa de diálogo. Seu código deve lidar com esse evento, conforme mostrado na próxima seção.

Considere um cenário no qual a interface do usuário da caixa de diálogo está relacionada à planilha ativa atualmente e à posição dessa planilha em relação às outras planilhas. No exemplo a seguir, `sheetPropertiesChanged` envia as propriedades da planilha do Excel para a caixa de diálogo. Nesse caso, a planilha atual se chama "Minha Planilha" e é a segunda folha na pasta de trabalho. Os dados são encapsulados em um objeto e stringizados para que possam ser passados para `messageChild`.

```javascript
function sheetPropertiesChanged() {
    const messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>Manipular DialogParentMessageReceived na caixa de diálogo

No JavaScript da caixa de diálogo, registre um manipulador para o `DialogParentMessageReceived` evento com o método [UI.addHandlerAsync](/javascript/api/office/office.ui#office-office-ui-addhandlerasync-member(1)) . Normalmente, isso é feito na [função Office.onReady ou Office.initialize](initialize-add-in.md), conforme mostrado no seguinte. (Um exemplo mais robusto é incluído posteriormente neste artigo.)

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

Em seguida, defina o `onMessageFromParent` manipulador. O código a seguir continua o exemplo da seção anterior. Observe que o Office passa um argumento para o manipulador e que a `message` propriedade do objeto argumento contém a cadeia de caracteres da página host. Neste exemplo, a mensagem é reconvertida para um objeto e jQuery é usada para definir o título superior da caixa de diálogo para corresponder ao novo nome da planilha.

```javascript
function onMessageFromParent(arg) {
    const messageFromParent = JSON.parse(arg.message);
    $('h1').text(messageFromParent.name);
}
```

É uma prática recomendada verificar se o manipulador está registrado corretamente. Você pode fazer isso passando um retorno de chamada para o `addHandlerAsync` método. Isso é executado quando a tentativa de registrar o manipulador é concluída. Use o manipulador para registrar ou mostrar um erro se o manipulador não tiver sido registrado com êxito. Apresentamos um exemplo a seguir. Observe que `reportError` é uma função, não definida aqui, que registra ou exibe o erro.

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

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a>Mensagens condicionais da página pai para a caixa de diálogo

Como você pode fazer várias `messageChild` chamadas da página host, mas tem apenas um manipulador na caixa de diálogo para o `DialogParentMessageReceived` evento, o manipulador deve usar a lógica condicional para distinguir mensagens diferentes. Você pode fazer isso de uma maneira exatamente paralela à forma como você estruturaria mensagens condicionais quando a caixa de diálogo está enviando uma mensagem para a página do host, conforme descrito em [mensagens condicionais](#conditional-messaging).

> [!NOTE]
> Em algumas situações, a `messageChild` API, que faz parte do conjunto de [requisitos DialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets), pode não ter suporte. Algumas maneiras alternativas para mensagens pai-a-caixa de diálogo são descritas em [formas alternativas de passar mensagens para uma caixa de diálogo de sua página de host](parent-to-dialog.md).

> [!IMPORTANT]
> O [conjunto de requisitos DialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets) não pode ser especificado na **\<Requirements\>** seção de um manifesto de suplemento. Você precisará verificar se há suporte para o DialogApi 1.2 no runtime usando o `isSetSupported` método conforme descrito em [Verificações do Runtime para o método e o suporte de conjunto de requisitos](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). O suporte para requisitos de manifesto está em desenvolvimento.

### <a name="cross-domain-messaging-to-the-dialog-runtime"></a>Mensagens entre domínios no runtime da caixa de diálogo

Depois que a caixa de diálogo for aberta, a caixa de diálogo ou o runtime pai poderão navegar para longe do domínio do suplemento. Se alguma dessas coisas acontecer, as chamadas para `messageChild` falharão, a menos que o código especifique o domínio do runtime da caixa de diálogo. Você faz isso adicionando um parâmetro [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) à chamada de `messageChild`. Esse objeto tem uma `targetOrigin` propriedade que especifica o domínio para o qual a mensagem deve ser enviada. Se o parâmetro não for usado, o Office pressupõe que o destino seja o mesmo domínio que o runtime pai está hospedando no momento.

> [!NOTE]
> Usar `messageChild` para enviar uma mensagem entre domínios requer o [conjunto de requisitos Origem do Diálogo 1.1](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets). O `DialogMessageOptions` parâmetro é ignorado em versões mais antigas do Office que não dão suporte ao conjunto de requisitos, portanto, o comportamento do método não será afetado se você passá-lo.

A seguir está um exemplo de uso `messageChild` para enviar uma mensagem entre domínios.

```js
dialog.messageChild(messageToDialog, { targetOrigin: "https://resource.contoso.com" });
```

Se a mensagem não incluir dados confidenciais, você poderá definir o `targetOrigin` como "\*" que permite que ele seja *enviado* para qualquer domínio. Apresentamos um exemplo a seguir.

```js
dialog.messageChild(messageToDialog, { targetOrigin: "*" });
```

Como o runtime que está hospedando a caixa de diálogo não pode acessar a **\<AppDomains\>** seção do manifesto e, assim, determinar se o domínio *do qual a mensagem vem* é confiável, você deve usar o `DialogParentMessageReceived` manipulador para determinar isso. O objeto que é passado para o manipulador contém o domínio que atualmente está hospedado no pai como sua `origin` propriedade. A seguir está um exemplo de como usar a propriedade.

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

Por exemplo, seu código pode usar a [função Office.onReady ou Office.initialize](initialize-add-in.md) para armazenar uma matriz de domínios confiáveis em uma variável global. Em `arg.origin` seguida, a propriedade pode ser verificada em relação a essa lista no manipulador.

> [!TIP]
> O `DialogMessageOptions` parâmetro foi adicionado ao `messageChild` método como um parâmetro necessário em meados de 2021. Os suplementos mais antigos que enviam uma mensagem entre domínios com o método não funcionam mais até que sejam atualizados para usar o novo parâmetro. Até que o suplemento seja atualizado, *somente no Office no Windows*, usuários e administradores do sistema podem permitir que esses suplementos continuem funcionando especificando os domínios confiáveis com uma configuração de registro: **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**. Para fazer isso, crie um arquivo com uma `.reg` extensão, salve-o no computador Windows e clique duas vezes nele para executá-lo. A seguir está um exemplo do conteúdo de tal arquivo.
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="close-the-dialog-box"></a>Fechar a caixa de diálogo

You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example.

```js
function closeButtonClick() {
    const messageObject = {messageType: "dialogClosed"};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

O manipulador de página host de `DialogMessageReceived` poderia chamar `dialog.close`, como neste exemplo. (Veja exemplos anteriores que mostram como o objeto `dialog` é inicializado.)

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
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

### <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>Usar a API de diálogo do Office com aplicativos de página única e roteamento do lado do cliente

SPAs e o roteamento do lado do cliente devem ser manuseados com cuidado ao usar a API de diálogo do Office. Confira [práticas recomendadas para usar o Office Dialog API em um SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).

### <a name="error-and-event-handling"></a>Manipulação de erros e eventos

Confira [Manipulando erros e eventos na caixa de diálogo do Office](dialog-handle-errors-events.md).

## <a name="next-steps"></a>Próximas etapas

Saiba mais sobre as armadilhas e as práticas recomendadas para a API de diálogo do Office em [práticas recomendadas e regras para a API do Office Dialog](dialog-best-practices.md).

## <a name="samples"></a>Exemplos

Todos os exemplos a seguir usam `displayDialogAsync`. Alguns têm servidores baseados em NodeJS e outros têm servidores ASP.NET/IIS-based, mas a lógica de usar o método é a mesma, independentemente de como o lado do servidor do suplemento é implementado.

**Básico:**

- [Exemplo da API da caixa de diálogo do suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [Conteúdo de treinamento/suplementos de construção (vários exemplos)](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

**Exemplos mais complexos:**

- [AsPNET do Microsoft Graph de Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Suplemento do Office Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [SSO do NodeJS do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)
- [SSO DO ASPNET de Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)
- [Exemplo de monetização SAAS do Suplemento do Office](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [AsPNET do Microsoft Graph do Suplemento do Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [SSO de suplemento do Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)
- [Visualizador de token de suplemento do Outlook](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Mensagem acionável do Suplemento do Outlook](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [Compartilhamento de suplementos do Outlook no OneDrive](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [Suplemento do PowerPoint Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Cenário de Runtime Compartilhado do Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario)
- [Guias Rápidos ASPNET do Suplemento do Excel](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Redact JS do Suplemento do Word](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Suplemento do Word JS SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [OAuth do cliente AngularJS do Suplemento do Word](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [Suplemento do Office Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [OAuth.io de suplemento do Office](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [Código de padrões de design de UX de suplemento do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

** Veja também**

- [Runtimes em suplementos do Office](../testing/runtimes.md)