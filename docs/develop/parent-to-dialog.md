---
title: Passando dados e mensagens para uma caixa de diálogo da página host
description: Saiba como transmitir dados para uma caixa de diálogo da página host usando as APIs messageChild e DialogParentMessageReceived.
ms.date: 04/16/2020
localization_priority: Normal
ms.openlocfilehash: cd332a58aa79a81aab7cf5a3d247950ce8bc655e
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547054"
---
# <a name="passing-data-and-messages-to-a-dialog-box-from-its-host-page-preview"></a>Passando dados e mensagens para uma caixa de diálogo da página de host (visualização)

O suplemento pode enviar mensagens da [página host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) para uma caixa de diálogo usando o método [MessageChild](/javascript/api/office/office.dialog#messagechild-message-) do objeto [Dialog](/javascript/api/office/office.dialog) .

> [!Important]
>
> - As APIs descritas neste artigo estão em visualização. Eles estão disponíveis para os desenvolvedores de experimentação; Mas não deve ser usado em um suplemento de produção. Até que esta API seja liberada, use as técnicas descritas em [passar informações para a caixa de diálogo](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) para suplementos de produção.
> - As APIs descritas neste artigo exigem o Office 365 (a versão de assinatura do Office). Você deve usar o build e a versão mensal mais recente do canal Insiders. É necessário ingressar no programa Office Insider para obter essa versão. Para saber mais, confira a página [Seja um Office Insider](https://insider.office.com). Observe que, quando uma compilação é graduada para o canal semestral de produção, o suporte para recursos de visualização é desativado para essa compilação.
> - Na fase inicial da visualização, as APIs têm suporte no Excel, PowerPoint e Word; Mas não no Outlook.
>
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## <a name="use-messagechild-from-the-host-page"></a>Usar `messageChild()` na página host

Quando você chama a API de diálogo do Office para abrir uma caixa de diálogo, um objeto [Dialog](/javascript/api/office/office.dialog) é retornado. Ele deve ser atribuído a uma variável, que geralmente tem escopo maior do que o método [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) porque o objeto será referenciado por outros métodos. Este é um exemplo:

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

Este `Dialog` objeto tem um método [messageChild](/javascript/api/office/office.dialog#messagechild-message-) que envia qualquer cadeia de caracteres ou em formato dados para a caixa de diálogo. Isso gera um `DialogParentMessageReceived` evento na caixa de diálogo. O código deve lidar com esse evento, conforme mostrado na próxima seção.

Considere um cenário em que a interface do usuário da caixa de diálogo deve se correlacionar com a planilha ativa no momento e a posição dessa planilha em relação às outras planilhas. No exemplo a seguir, `sheetPropertiesChanged` envia as propriedades de planilha do Excel para a caixa de diálogo. Nesse caso, a planilha atual é chamada "minha planilha" e é a 2ª folha na pasta de trabalho. Os dados são encapsulados em um objeto que é em formato para que ele possa ser passado para `messageChild`.

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

## <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>Manipular DialogParentMessageReceived na caixa de diálogo

No JavaScript da caixa de diálogo, registre um manipulador para o `DialogParentMessageReceived` evento com o método [UI. addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) . Isso geralmente é feito nos [métodos Office. onReady ou Office. Initialize](initialize-add-in.md). Este é um exemplo:

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

Em seguida, defina `onMessageFromParent` o manipulador. O código a seguir continua o exemplo da seção anterior. Observe que o Office passa um argumento para o manipulador e que `message` a propriedade do objeto Argument contém a cadeia de caracteres da página host. Neste exemplo, a mensagem é convertida para um objeto e o jQuery é usado para definir o título superior da caixa de diálogo para corresponder ao novo nome da planilha.

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

É uma prática recomendada verificar se o manipulador está registrado corretamente. Você pode fazer isso passando um retorno de chamada para `addHandlerAsync` o método que é executado quando a tentativa de registrar o manipulador é concluída. Use o manipulador para registrar ou mostrar um erro se o manipulador não tiver sido registrado com êxito. Apresentamos um exemplo a seguir. Observe que `reportError` é uma função, não definida aqui, que registra ou exibe o erro.

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

## <a name="conditional-messaging"></a>Mensagens condicionais

Como você pode fazer várias `messageChild` chamadas a partir da página host, mas tem apenas um manipulador na caixa de diálogo para o `DialogParentMessageReceived` evento, o manipulador deve usar a lógica condicional para distinguir mensagens diferentes. Você pode fazer isso de uma maneira que seja precisamente paralela à forma como você estruturaria mensagens condicionais quando a caixa de diálogo estiver enviando uma mensagem para a página host, conforme descrito em [mensagens condicionais](dialog-api-in-office-add-ins.md#conditional-messaging).
