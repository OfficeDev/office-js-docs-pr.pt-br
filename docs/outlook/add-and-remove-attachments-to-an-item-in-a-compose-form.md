---
title: Adicionar e remover os anexos em um suplemento do Outlook
description: Você pode usar várias APIs de anexo para gerenciar os arquivos ou os itens do Outlook anexados ao item que o usuário está redigindo.
ms.date: 11/11/2020
localization_priority: Normal
ms.openlocfilehash: 6f146b3efc3234313191d93af05d9c0d35111829
ms.sourcegitcommit: 5bfd1e9956485c140179dfcc9d210c4c5a49a789
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/13/2020
ms.locfileid: "49071701"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a>Gerenciar anexos de um item em um formulário de composição no Outlook

A API JavaScript do Office fornece várias APIs que você pode usar para gerenciar anexos de um item quando o usuário está redigindo.

## <a name="attach-a-file-or-outlook-item"></a>Anexar um arquivo ou item do Outlook

Você pode anexar um arquivo ou item do Outlook a um formulário de composição usando o método apropriado para o tipo de anexo.

- [addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): anexar um arquivo
- [addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): anexar um arquivo usando sua cadeia de caracteres Base64
- [addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): anexar um item do Outlook

Estes são métodos assíncronos, o que significa que a execução pode prosseguir sem esperar que a ação seja concluída. Dependendo do local original e do tamanho do anexo que está sendo adicionado, a chamada assíncrona poderá levar algum tempo para ser concluída.

Se houver tarefas que dependam da conclusão da ação, você deverá executá-las em um método de retorno de chamada. Esse método de retorno de chamada é opcional e é invocado quando o carregamento do anexo é concluído. O método de retorno de chamada usa um objeto [AsyncResult](/javascript/api/office/office.asyncresult) como um parâmetro de saída que fornece qualquer status, erro e valor retornado da adição do anexo. Se o retorno de chamada requer parâmetros adicionais, você pode especificá-los no parâmetro opcional `options.asyncContext`. `options.asyncContext` pode ser de qualquer tipo que seu método de retorno de chamada espere.

Por exemplo, você pode definir `options.asyncContext` como um objeto JSON que contém um ou mais pares chave-valor. Você pode encontrar mais exemplos sobre como passar parâmetros opcionais para métodos assíncronos na plataforma de suplementos do Office em [Programação assíncrona em suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). O exemplo a seguir mostra como usar o parâmetro `asyncContext` para passar dois argumentos para um método de retorno de chamada:

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

Você pode verificar o sucesso ou o erro de uma chamada de método assíncrono no método de retorno de chamada usando as propriedades `status` e `error` do objeto `AsyncResult`. Se a anexação for concluída com êxito, você poderá usar a propriedade `AsyncResult.value` para obter a ID do anexo. A ID do anexo é um número inteiro que você pode usar posteriormente para remover o anexo.

> [!NOTE]
> Como prática recomendada, você só deverá usar a ID do anexo para remover um anexo se o mesmo suplemento tiver adicionado esse anexo na mesma sessão. No Outlook na Web e dispositivos móveis, a ID do anexo é válida apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o suplemento ou se o usuário começa a redigir em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.

### <a name="attach-a-file"></a>Anexar um arquivo

Você pode anexar um arquivo a uma mensagem ou compromisso em um formulário de composição usando o `addFileAttachmentAsync` método e especificando o URI do arquivo. Você também pode usar o `addFileAttachmentFromBase64Async` método, mas especificar a cadeia de caracteres Base64 como entrada. Se o arquivo estiver protegido, você poderá incluir uma identidade ou um token de autenticação apropriado como um parâmetro de cadeia de caracteres de consulta de URI. O Exchange fará uma chamada à URI para obter o anexo, e o serviço Web que protege o arquivo precisará usar o token como um meio de autenticação.

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
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
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

### <a name="attach-an-outlook-item"></a>Anexar um item do Outlook

Você pode anexar um item do Outlook (por exemplo, um item de email, calendário ou contato) a uma mensagem ou a um compromisso em um formulário de redação, especificando a ID do item dos EWS (Serviços Web do Exchange) e usando o método `addItemAttachmentAsync`. Você pode obter a ID dos EWS de um item de email, calendário, contato ou tarefa na caixa de correio do usuário, usando o método [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) e acessando a operação [FindItem](/exchange/client-developer/web-service-reference/finditem-operation) do EWS. A propriedade [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) também fornece a ID dos EWS de um item existente em um formulário de leitura.

A seguinte função JavaScript, `addItemAttachment` , estende o primeiro exemplo acima e adiciona um item como um anexo ao email ou compromisso que está sendo composto. A função utiliza como argumento a ID dos EWS do item que será anexado. Se a conexão for bem-sucedida, ela receberá a ID de anexo para processamento adicional, incluindo a remoção desse anexo na mesma sessão.

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
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                var attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> Você pode usar um suplemento de redação para anexar uma instância de um compromisso recorrente no Outlook na Web ou em dispositivos móveis. No entanto, em um cliente avançado do Outlook com suporte, tentar anexar uma instância resultaria na anexação da série recorrente (o compromisso mestre).

## <a name="get-attachments"></a>Obter anexos

As APIs para obter anexos no modo de composição estão disponíveis no [conjunto de requisitos 1,8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).

- [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

Você pode usar o método [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) para obter os anexos da mensagem ou do compromisso que está sendo composto.

Para obter o conteúdo de um anexo, você pode usar o método [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) . Os formatos suportados estão listados na enumeração [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) .

Você deve fornecer um método de retorno de chamada para verificar o status e o erro usando o `AsyncResult` objeto Parameter de saída. Você também pode passar qualquer parâmetro adicional para o método de retorno de chamada usando o `asyncContext` parâmetro Optional.

O seguinte exemplo de JavaScript Obtém os anexos e permite que você configure a manipulação distinta para cada formato de anexo com suporte.

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

Você pode remover um anexo de arquivo ou item de um item de mensagem ou compromisso em um formulário de composição especificando a ID correspondente do anexo e usando o método [removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods). Você só deve remover os anexos que o mesmo suplemento adicionou na mesma sessão. Semelhante aos `addFileAttachmentAsync` métodos e `addItemAttachmentAsync` , `removeAttachmentAsync` é um método assíncrono. Você deve fornecer um método de retorno de chamada para verificar o status e o erro usando o `AsyncResult` objeto Parameter de saída. Você também pode passar qualquer parâmetro adicional para o método de retorno de chamada usando o `asyncContext` parâmetro Optional.

A seguinte função JavaScript, `removeAttachment` , continua a estender os exemplos acima e remove o anexo especificado do email ou compromisso que está sendo composto. A função utiliza como argumento a ID do anexo a ser removido. Você pode obter a ID de um anexo após uma chamada bem-sucedida, `addFileAttachmentAsync` `addFileAttachmentFromBase64Async` ou de `addItemAttachmentAsync` método, e armazená-lo para uma `removeAttachmentAsync` chamada de método subsequente.

```js
// Removes the specified attachment from the composed item.
// ID is the Exchange identifier of the attachment to be
// removed.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the
    // callback method is invoked. Here, the callback
    // method uses an asyncResult parameter and gets
    // the ID of the removed attachment if the removal
    // succeeds.
    // You can optionally pass any object you wish to
    // access in the callback method as an argument to
    // the asyncContext parameter.
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
