---
title: Adicionar e remover os anexos em um suplemento do Outlook
description: Você pode usar várias APIs de anexo para gerenciar os arquivos ou Outlook itens anexados ao item que o usuário está compondo.
ms.date: 02/24/2021
localization_priority: Normal
ms.openlocfilehash: 0ba142bb1e8fb5f324d2bb6460bc8325a4800d2d
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348584"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a>Gerenciar anexos de um item em um formulário de redação em Outlook

A Office api JavaScript oferece várias APIs que você pode usar para gerenciar anexos de um item quando o usuário está compondo.

## <a name="attach-a-file-or-outlook-item"></a>Anexar um arquivo ou Outlook item

Você pode anexar um arquivo ou Outlook item a um formulário de composição usando o método apropriado para o tipo de anexo.

- [addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): anexar um arquivo
- [addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): anexar um arquivo usando sua cadeia de caracteres base64
- [addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): anexar um Outlook item

Esses são métodos assíncronos, o que significa que a execução pode continuar sem aguardar a conclusão da ação. Dependendo do local original e do tamanho do anexo que está sendo adicionado, a chamada assíncrona pode demorar um pouco para ser concluída.

Se houver tarefas que dependam da conclusão da ação, você deverá executá-las em um método de retorno de chamada. Esse método de retorno de chamada é opcional e é invocado quando o carregamento do anexo é concluído. O método de retorno de chamada usa um objeto [AsyncResult](/javascript/api/office/office.asyncresult) como um parâmetro de saída que fornece qualquer status, erro e valor retornado da adição do anexo. Se o retorno de chamada requer parâmetros adicionais, você pode especificá-los no parâmetro opcional `options.asyncContext`. `options.asyncContext` pode ser de qualquer tipo que seu método de retorno de chamada espere.

Por exemplo, você pode definir `options.asyncContext` como um objeto JSON que contém um ou mais pares de valores-chave. Você pode encontrar mais exemplos sobre a passagem de parâmetros opcionais para métodos assíncronos na plataforma de Office Add-ins na programação [assíncrona em Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). O exemplo a seguir mostra como usar o `asyncContext` parâmetro para passar 2 argumentos para um método de retorno de chamada.

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

Você pode verificar o sucesso ou o erro de uma chamada de método assíncrono no método de retorno de chamada usando as propriedades `status` e `error` do objeto `AsyncResult`. Se a anexação for concluída com êxito, você poderá usar a propriedade `AsyncResult.value` para obter a ID do anexo. A ID do anexo é um número inteiro que você pode usar posteriormente para remover o anexo.

> [!NOTE]
> A ID do anexo é válida somente dentro da mesma sessão e não tem garantia de mapear para o mesmo anexo entre as sessões. Os exemplos de quando uma sessão acabou incluem quando o usuário fecha o complemento ou se o usuário começa a compor em um formulário em linha e, subsequentemente, sai do formulário em linha para continuar em uma janela separada.

### <a name="attach-a-file"></a>Anexar um arquivo

Você pode anexar um arquivo a uma mensagem ou compromisso em um formulário de composição usando o método e especificando o `addFileAttachmentAsync` URI do arquivo. Você também pode usar o `addFileAttachmentFromBase64Async` método, mas especificar a cadeia de caracteres base64 como entrada. Se o arquivo estiver protegido, você poderá incluir uma identidade ou um token de autenticação apropriado como um parâmetro de cadeia de caracteres de consulta de URI. O Exchange fará uma chamada à URI para obter o anexo, e o serviço Web que protege o arquivo precisará usar o token como um meio de autenticação.

O exemplo de JavaScript a seguir é um suplemento de redação que anexa um arquivo, picture.png, de um servidor Web à mensagem ou ao compromisso que está sendo redigido. O método de retorno de chamada usa `asyncResult` como um parâmetro, verifica o status de resultado e obtém a ID do anexo caso o método tenha êxito.

```js
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback method is invoked and gets the attachment ID.
        // You can optionally pass any object that you would
        // access in the callback method as an argument to
        // the asyncContext parameter.
        Office.context.mailbox.item.addFileAttachmentAsync(
            `https://webserver/picture.png`,
            'picture.png',
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                } else {
                    // Get the ID of the attached file.
                    var attachmentID = asyncResult.value;
                    write('ID of added attachment: ' + attachmentID);
                }
            });
    });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="attach-an-outlook-item"></a>Anexar um Outlook de Outlook

Você pode anexar um item de Outlook (por exemplo, email, calendário ou item de contato) a uma mensagem ou compromisso em um formulário de redação especificando a ID do Exchange Web Services (EWS) do item e usando o `addItemAttachmentAsync` método. Você pode obter a ID do EWS de um email, calendário, contato ou item de tarefa na caixa de correio do usuário usando o método [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) e acessando a operação EWS [FindItem](/exchange/client-developer/web-service-reference/finditem-operation). A propriedade [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) também fornece a ID dos EWS de um item existente em um formulário de leitura.

A função JavaScript a seguir, , estende o primeiro exemplo acima e adiciona um item como um anexo ao email ou compromisso `addItemAttachment` que está sendo composto. A função utiliza como argumento a ID dos EWS do item que será anexado. Se a anexação for bem-sucedida, ela obtém a ID do anexo para processamento posterior, incluindo a remoção desse anexo na mesma sessão.

```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(itemId) {
    // When the attachment finishes uploading, the
    // callback method is invoked. Here, the callback
    // method uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to
    // access in the callback method as an argument to
    // the asyncContext parameter.
    Office.context.mailbox.item.addItemAttachmentAsync(
        itemId,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            } else {
                var attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> Você pode usar um complemento de composição para anexar uma instância de um compromisso recorrente em Outlook na Web ou em dispositivos móveis. No entanto, em um cliente Outlook desktop de suporte, tentar anexar uma instância resultaria em anexar a série recorrente (o compromisso pai).

## <a name="get-attachments"></a>Obter anexos

APIs para obter anexos no modo de redação estão disponíveis no conjunto [de requisitos 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).

- [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

Você pode usar o [método getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) para obter os anexos da mensagem ou do compromisso que está sendo composto.

Para obter o conteúdo de um anexo, você pode usar o [método getAttachmentContentAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) Os formatos com suporte estão listados na [enumeração AttachmentContentFormat.](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

Você deve fornecer um método de retorno de chamada para verificar o status e qualquer erro usando o `AsyncResult` objeto de parâmetro de saída. Você também pode passar quaisquer parâmetros adicionais para o método de retorno de chamada usando o parâmetro `asyncContext` opcional.

O exemplo JavaScript a seguir obtém os anexos e permite configurar o tratamento distinto para cada formato de anexo com suporte.

```js
var item = Office.context.mailbox.item;
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

## <a name="remove-an-attachment"></a>Remover um anexo

Você pode remover um anexo de arquivo ou item de uma mensagem ou item de compromisso em um formulário de composição especificando a ID de anexo correspondente ao usar o [método removeAttachmentAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

> [!IMPORTANT]
> Se você estiver usando o conjunto de requisitos 1.7 ou anterior, você deve remover apenas anexos que o mesmo complemento adicionou na mesma sessão.

Semelhante aos `addFileAttachmentAsync` métodos , e é um método `addItemAttachmentAsync` `getAttachmentsAsync` `removeAttachmentAsync` assíncrono. Você deve fornecer um método de retorno de chamada para verificar o status e qualquer erro usando o `AsyncResult` objeto de parâmetro de saída. Você também pode passar quaisquer parâmetros adicionais para o método de retorno de chamada usando o parâmetro `asyncContext` opcional.

A função JavaScript a seguir, , continua a estender os exemplos acima e remove o anexo especificado do email ou compromisso `removeAttachment` que está sendo composto. A função utiliza como argumento a ID do anexo a ser removido. Você pode obter a ID de um anexo após uma chamada bem-sucedida , ou método, e `addFileAttachmentAsync` usá-lo em uma chamada de método `addFileAttachmentFromBase64Async` `addItemAttachmentAsync` `removeAttachmentAsync` subsequente. Você também pode chamar (introduzido no conjunto de requisitos 1.8) para obter os anexos e suas IDs para essa sessão `getAttachmentsAsync` de complemento.

```js
// Removes the specified attachment from the composed item.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the callback method is invoked.
    // Here, the callback method uses an asyncResult parameter and
    // gets the ID of the removed attachment if the removal succeeds.
    // You can optionally pass any object you wish to access in the
    // callback method as an argument to the asyncContext parameter.
    Office.context.mailbox.item.removeAttachmentAsync(
        attachmentId,
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            } else {
                write('Removed attachment with the ID: ' + asyncResult.value);
            }
        });
}
```

## <a name="see-also"></a>Confira também

- [Criar suplementos do Outlook para formulários de redação](compose-scenario.md)
- [Programação assíncrona em Suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md)
