---
title: Manipulando erros e eventos na caixa de diálogo do Office
description: Descreve como capturar e lidar com erros ao abrir e usar a caixa de diálogo do Office
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: d83d5c4627f68c3f4b1c196cf543d01bf981abbe
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608171"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a>Manipulando erros e eventos na caixa de diálogo do Office

Este artigo descreve como capturar e lidar com erros ao abrir a caixa de diálogo e os erros que ocorrem dentro da caixa de diálogo.

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com as noções básicas de usar a API de caixa de diálogo do Office, conforme descrito em [usar a API de caixa de diálogo do Office em seus suplementos do Office](dialog-api-in-office-add-ins.md).
> 
> Consulte também [práticas recomendadas e regras para a API de diálogo do Office](dialog-best-practices.md).

Seu código deve manipular duas categorias de eventos:

- Erros retornados pela chamada de `displayDialogAsync` porque não foi possível criar a caixa de diálogo.
- Erros e outros eventos, na caixa de diálogo.

## <a name="errors-from-displaydialogasync"></a>Erros de displayDialogAsync

Além dos erros gerais de plataforma e de sistema, quatro erros são específicos para chamar `displayDialogAsync` .

|Número do código|Significado|
|:-----|:-----|
|12004|O domínio que a URL transmitiu para `displayDialogAsync` não é confiável. O domínio deve ser o mesmo domínio que o da página de host (incluindo o protocolo e o número de porta).|
|12005|A URL passada para `displayDialogAsync` usa o protocolo HTTP. HTTPS é necessário. (Em algumas versões do Office, o texto da mensagem de erro retornado com 12005 é o mesmo retornado para 12004.)|
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|Uma caixa de diálogo já está aberta na janela do host. Uma janela do host, como um painel de tarefas, só pode ter uma caixa de diálogo aberta por vez.|
|12009|O usuário opta por ignorar a caixa de diálogo. Este erro pode ocorrer no Office na Web, onde os usuários podem optar por não permitir que um suplemento apresente uma caixa de diálogo. Para obter mais informações, consulte [lidando de bloqueadores de pop-up com o Office na Web](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).|

Quando `displayDialogAsync` é chamado, ele passa um objeto [AsyncResult](/javascript/api/office/office.asyncresult) para sua função de retorno de chamada. Quando a chamada for bem-sucedida, a caixa de diálogo será aberta e a `value` Propriedade do `AsyncResult` objeto será um objeto [Dialog](/javascript/api/office/office.dialog) . Para obter um exemplo disso, consulte [enviar informações da caixa de diálogo para a página host](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page). Quando a chamada `displayDialogAsync` falhar, a caixa de diálogo não é criada, a `status` Propriedade do `AsyncResult` objeto é definida como `Office.AsyncResultStatus.Failed` e a `error` Propriedade do objeto é preenchida. Você sempre deve fornecer um retorno de chamada que testa o `status` e responde quando é um erro. Para obter um exemplo que relata a mensagem de erro independentemente de seu número de código, consulte o código a seguir. (A `showNotification` função, não definida neste artigo, exibe ou registra o erro. Para obter um exemplo de como você pode implementar essa função no seu suplemento, confira [exemplo de API de caixa de diálogo do suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)

```js
var dialog;
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

Três erros e eventos na caixa de diálogo irão gerar um `DialogEventReceived` evento na página host. Para obter um lembrete sobre o que é uma página de host, consulte [abrir uma caixa de diálogo em uma página de host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).

|Número do código|Significado|
|:-----|:-----|
|12002|Uma destas opções:<br> - Não existe uma página na URL transmitida para `displayDialogAsync`.<br> – A página que foi passada para `displayDialogAsync` carregado, mas a caixa de diálogo foi redirecionada para uma página que não pode ser encontrada ou carregada, ou foi direcionada para uma URL com sintaxe inválida.|
|12003|A caixa de diálogo foi direcionada para uma URL com o protocolo HTTP. HTTPS é necessário.|
|12006|A caixa de diálogo foi fechada, geralmente porque o usuário escolheu o botão **fechar** **X**.|

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
