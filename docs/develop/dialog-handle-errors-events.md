---
title: Manipulando erros e eventos na caixa de diálogo do Office
description: Saiba como interceptar e manipular erros ao abrir e usar a caixa de diálogo do Office.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0e8eefe4ee868a3cdc52ee8d425271435404bc04
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889454"
---
# <a name="handle-errors-and-events-in-the-office-dialog-box"></a>Manipular erros e eventos na caixa de diálogo do Office

Este artigo descreve como interceptar e manipular erros ao abrir a caixa de diálogo e erros que ocorrem dentro da caixa de diálogo.

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com os conceitos básicos do uso da API de caixa de diálogo do Office, conforme descrito em Usar a API de diálogo do Office em seus [Suplementos do Office](dialog-api-in-office-add-ins.md).
>
> Confira também [as práticas recomendadas e as regras para a API de caixa de diálogo do Office](dialog-best-practices.md).

Seu código deve lidar com duas categorias de eventos.

- Erros retornados pela chamada de `displayDialogAsync` porque não foi possível criar a caixa de diálogo.
- Erros e outros eventos na caixa de diálogo.

## <a name="errors-from-displaydialogasync"></a>Erros de displayDialogAsync

Além dos erros gerais da plataforma e do sistema, quatro erros são específicos da chamada `displayDialogAsync`.

|Número do código|Significado|
|:-----|:-----|
|12004|O domínio que a URL transmitiu para `displayDialogAsync` não é confiável. O domínio deve ser o mesmo domínio que o da página de host (incluindo o protocolo e o número de porta).|
|12005|A URL passada para `displayDialogAsync` usa o protocolo HTTP. HTTPS é necessário. (Em algumas versões do Office, o texto da mensagem de erro retornado com 12005 é o mesmo retornado para 12004.)|
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|Uma caixa de diálogo já está aberta na janela do host. Uma janela do host, como um painel de tarefas, só pode ter uma caixa de diálogo aberta por vez.|
|12009|O usuário opta por ignorar a caixa de diálogo. Esse erro pode ocorrer no Office na Web, em que os usuários podem optar por não permitir que um suplemento apresente uma caixa de diálogo. Para obter mais informações, consulte [Tratamento de bloqueadores pop-up com Office na Web](dialog-best-practices.md#handle-pop-up-blockers-with-office-on-the-web).|

Quando `displayDialogAsync` é chamado, ele passa um [objeto AsyncResult](/javascript/api/office/office.asyncresult) para sua função de retorno de chamada. Quando a chamada é bem-sucedida, a caixa de diálogo é aberta e `value` a propriedade do objeto `AsyncResult` é um [objeto Dialog](/javascript/api/office/office.dialog) . Para obter um exemplo disso, consulte [Enviar informações da caixa de diálogo para a página host](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page). Quando a chamada para `displayDialogAsync` falhar, a caixa de diálogo não é criada, `status` `AsyncResult` `Office.AsyncResultStatus.Failed`a propriedade do objeto é definida como e `error` a propriedade do objeto é preenchida. Você sempre deve fornecer um retorno de chamada que testa `status` e responde quando é um erro. Para obter um exemplo que relata a mensagem de erro, independentemente de seu número de código, consulte o código a seguir. (A `showNotification` função, não definida neste artigo, exibe ou registra o erro. Para obter um exemplo de como você pode implementar essa função em seu suplemento, consulte o Exemplo de API de Caixa de Diálogo [do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)

```js
let dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

## <a name="errors-and-events-in-the-dialog-box"></a>Erros e eventos na caixa de diálogo

Três erros e eventos na caixa de diálogo gerarão um `DialogEventReceived` evento na página host. Para obter um lembrete do que é uma página host, [consulte Abrir uma caixa de diálogo de uma página host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).

|Número do código|Significado|
|:-----|:-----|
|12002|Uma destas opções:<br> - Não existe uma página na URL transmitida para `displayDialogAsync`.<br> - A página `displayDialogAsync` que foi passada para carregada, mas a caixa de diálogo foi então redirecionada para uma página que não pode localizar ou carregar, ou foi direcionada para uma URL com sintaxe inválida.|
|12003|A caixa de diálogo foi direcionada para uma URL com o protocolo HTTP. HTTPS é necessário.|
|12006|A caixa de diálogo foi fechada, geralmente porque o usuário escolheu o **botão** **Fechar X**.|

Seu código pode atribuir um manipulador para o evento `DialogEventReceived` na chamada para `displayDialogAsync`. Apresentamos um exemplo simples a seguir.

```js
let dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

Para obter um exemplo de um manipulador para o evento `DialogEventReceived` que cria mensagens de erro personalizadas para cada código de erro, consulte o exemplo a seguir.

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

## <a name="see-also"></a>Confira também

Para ver um suplemento de exemplo que manipula erros dessa forma, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).
