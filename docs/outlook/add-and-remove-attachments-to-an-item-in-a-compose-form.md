---
title: Adicionar e remover os anexos em um suplemento do Outlook
description: Use várias APIs de anexo para gerenciar os arquivos ou itens do Outlook anexados ao item que o usuário está redigindo.
ms.date: 08/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: af3b44814fd11c5e2006dbb921130c15c7535385
ms.sourcegitcommit: 76b8c79cba707c771ae25df57df14b6445f9b8fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2022
ms.locfileid: "67274166"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a>Gerenciar anexos de um item em um formulário de composição no Outlook

A API JavaScript do Office fornece várias APIs que você pode usar para gerenciar anexos de um item quando o usuário está redigindo.

## <a name="attach-a-file-or-outlook-item"></a>Anexar um arquivo ou item do Outlook

Você pode anexar um arquivo ou item do Outlook a um formulário de composição usando o método apropriado para o tipo de anexo.

- [addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods): anexar um arquivo
- [addFileAttachmentFromBase64Async](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods): anexar um arquivo usando sua cadeia de caracteres base64
- [addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods): anexar um item do Outlook

Esses são métodos assíncronos, o que significa que a execução pode continuar sem aguardar a conclusão da ação. Dependendo do local original e do tamanho do anexo que está sendo adicionado, a chamada assíncrona pode levar algum tempo para ser concluída.

Se houver tarefas que dependem da ação a ser concluída, você deverá executar essas tarefas em uma função de retorno de chamada. Essa função de retorno de chamada é opcional e é invocada quando o carregamento do anexo é concluído. A função de retorno de chamada usa um [objeto AsyncResult](/javascript/api/office/office.asyncresult) como um parâmetro de saída que fornece qualquer status, erro e valor retornado da adição do anexo. Se o retorno de chamada requer parâmetros adicionais, você pode especificá-los no parâmetro opcional `options.asyncContext`. `options.asyncContext` pode ser de qualquer tipo que sua função de retorno de chamada espera.

Por exemplo, você pode definir como `options.asyncContext` um objeto JSON que contém um ou mais pares chave-valor. Você pode encontrar mais exemplos sobre como passar parâmetros opcionais para métodos assíncronos na plataforma de Suplementos do Office na programação assíncrona em [Suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-to-asynchronous-methods). O exemplo a seguir mostra como usar o parâmetro `asyncContext` para passar dois argumentos para uma função de retorno de chamada.

```js
const options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

Você pode verificar se há êxito ou erro de uma chamada de método assíncrono `status` na função de retorno de chamada usando as propriedades `error` e o objeto `AsyncResult` . Se a anexação for concluída com êxito, você poderá usar `AsyncResult.value` a propriedade para obter a ID do anexo. A ID do anexo é um número inteiro que você pode usar posteriormente para remover o anexo.

> [!NOTE]
> A ID do anexo é válida somente dentro da mesma sessão e não há garantia de mapear para o mesmo anexo entre as sessões. Exemplos de quando uma sessão acabou incluem quando o usuário fecha o suplemento ou se o usuário começa a redigir em um formulário embutido e, subsequentemente, exibe o formulário embutido para continuar em uma janela separada.

### <a name="attach-a-file"></a>Anexar um arquivo

Você pode anexar um arquivo a `addFileAttachmentAsync` uma mensagem ou compromisso em um formulário de composição usando o método e especificando o URI do arquivo. Você também pode usar o método `addFileAttachmentFromBase64Async` , mas especificar a cadeia de caracteres base64 como entrada. Se o arquivo estiver protegido, você poderá incluir uma identidade ou um token de autenticação apropriado como um parâmetro de cadeia de caracteres de consulta de URI. O Exchange fará uma chamada à URI para obter o anexo, e o serviço Web que protege o arquivo precisará usar o token como um meio de autenticação.

O exemplo de JavaScript a seguir é um suplemento de composição que anexa um arquivo, picture.png, de um servidor Web à mensagem ou ao compromisso que está sendo redigido. A função de retorno de chamada `asyncResult` usa como parâmetro, verifica o status do resultado e obtém a ID do anexo se o método for bem-sucedido.

```js
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback function is invoked and gets the attachment ID.
        // You can optionally pass any object that you would
        // access in the callback function as an argument to
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
                    const attachmentID = asyncResult.value;
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

Para adicionar uma imagem base64 embutida ao corpo de uma mensagem ou compromisso que está sendo composto, primeiro você deve obter o corpo do item `Office.context.mailbox.item.body.getAsync` `addFileAttachmentFromBase64Async` atual usando o método antes de inserir a imagem usando o método. Caso contrário, a imagem não será renderizada no corpo depois de inserida. Para obter diretrizes, consulte o exemplo de JavaScript a seguir, que adiciona uma imagem base64 embutida ao início de um corpo de item.

```js
const mailItem = Office.context.mailbox.item;
const base64String =
  "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAnUExURQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAN0S+bUAAAAMdFJOUwAQIDBAUI+fr7/P7yEupu8AAAAJcEhZcwAADsMAAA7DAcdvqGQAAAF8SURBVGhD7dfLdoMwDEVR6Cspzf9/b20QYOthS5Zn0Z2kVdY6O2WULrFYLBaLxd5ur4mDZD14b8ogWS/dtxV+dmx9ysA2QUj9TQRWv5D7HyKwuIW9n0vc8tkpHP0W4BOg3wQ8wtlvA+PC1e8Ao8Ld7wFjQtHvAiNC2e8DdqHqKwCrUPc1gE1AfRVgEXBfB+gF0lcCWoH2tYBOYPpqQCNwfT3QF9i+AegJfN8CtAWhbwJagtS3AbIg9o2AJMh9M5C+SVGBvx6zAfmT0r+Bv8JMwP4kyFPir+cswF5KL3WLv14zAFBCLf56Tw9cparFX4upgaJUtPhrOS1QlY5W+vWTXrGgBFB/b72ev3/0igUdQPppP/nfowfKUUEFcP207y/yxKmgAYQ+PywoAFOfCH3A2MdCFzD3kdADBvq10AGG+pXQBgb7pdAEhvuF0AIc/VtoAK7+JciAs38KIuDugyAC/v4hiMCE/i7IwLRBsh68N2WQjMVisVgs9i5bln8LGScNcCrONQAAAABJRU5ErkJggg==";

// Get the current body of the message or appointment.
mailItem.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
  if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
    // Insert the base64 image to the beginning of the body.
    const options = { isInline: true, asyncContext: bodyResult.value };
    mailItem.addFileAttachmentFromBase64Async(base64String, "sample.png", options, (attachResult) => {
      if (attachResult.status === Office.AsyncResultStatus.Succeeded) {
        let body = attachResult.asyncContext;
        body = body.replace("<p class=MsoNormal>", `<p class=MsoNormal><img src="cid:sample.png">`);
        mailItem.body.setAsync(body, { coercionType: Office.CoercionType.Html }, (setResult) => {
          if (setResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Inline base64 image added to the body.");
          } else {
            console.log(setResult.error.message);
          }
        });
      } else {
        console.log(attachResult.error.message);
      }
    });
  } else {
    console.log(bodyResult.error.message);
  }
});
```

### <a name="attach-an-outlook-item"></a>Anexar um item do Outlook

Você pode anexar um item do Outlook (por exemplo, email, calendário ou item de contato) a uma mensagem ou compromisso em um formulário de composição especificando a ID dos Serviços Web do Exchange (EWS) do item `addItemAttachmentAsync` e usando o método. Você pode obter a ID do EWS de um email, calendário, contato ou item de tarefa na caixa de correio do usuário usando o método [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) e acessando a operação [EWS FindItem](/exchange/client-developer/web-service-reference/finditem-operation). A propriedade [item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) também fornece a ID dos EWS de um item existente em um formulário de leitura.

A função JavaScript a seguir estende `addItemAttachment`o primeiro exemplo acima e adiciona um item como um anexo ao email ou compromisso que está sendo composto. A função utiliza como argumento a ID dos EWS do item que será anexado. Se a anexação for bem-sucedida, ela obterá a ID do anexo para processamento adicional, incluindo a remoção desse anexo na mesma sessão.

```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(itemId) {
    // When the attachment finishes uploading, the
    // callback function is invoked. Here, the callback
    // function uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to
    // access in the callback function as an argument to
    // the asyncContext parameter.
    Office.context.mailbox.item.addItemAttachmentAsync(
        itemId,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            } else {
                const attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> Você pode usar um suplemento de composição para anexar uma instância de um compromisso recorrente no Outlook na Web ou em dispositivos móveis. No entanto, em um cliente de área de trabalho do Outlook com suporte, tentar anexar uma instância resultaria na anexação da série recorrente (o compromisso pai).

## <a name="get-attachments"></a>Obter anexos

As APIs para obter anexos no modo de composição estão disponíveis no conjunto [de requisitos 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8).

- [getAttachmentsAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)

Você pode usar o [método getAttachmentsAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) para obter os anexos da mensagem ou compromisso que está sendo composto.

Para obter o conteúdo de um anexo, você pode usar o [método getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) . Os formatos com suporte são listados na [enumeração AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) .

Você deve fornecer uma função de retorno de chamada para verificar o status e qualquer erro usando o objeto `AsyncResult` de parâmetro de saída. Você também pode passar quaisquer parâmetros adicionais para a função de retorno de chamada usando o parâmetro `asyncContext` opcional.

O exemplo de JavaScript a seguir obtém os anexos e permite que você configure a manipulação distinta para cada formato de anexo com suporte.

```js
const item = Office.context.mailbox.item;
const options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (let i = 0 ; i < result.value.length ; i++) {
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

Você pode remover um anexo de arquivo ou item de uma mensagem ou item de compromisso em um formulário de composição especificando a ID de anexo correspondente ao usar o [método removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) .

> [!IMPORTANT]
> Se você estiver usando o conjunto de requisitos 1.7 ou anterior, só deverá remover anexos que o mesmo suplemento adicionou na mesma sessão.

Semelhante a `addFileAttachmentAsync`, `addItemAttachmentAsync`e métodos `getAttachmentsAsync` , `removeAttachmentAsync` é um método assíncrono. Você deve fornecer uma função de retorno de chamada para verificar o status e qualquer erro usando o objeto `AsyncResult` de parâmetro de saída. Você também pode passar quaisquer parâmetros adicionais para a função de retorno de chamada usando o parâmetro `asyncContext` opcional.

A função JavaScript a seguir, `removeAttachment`continua estendendo os exemplos acima e remove o anexo especificado do email ou compromisso que está sendo composto. A função utiliza como argumento a ID do anexo a ser removido. Você pode obter a ID de um anexo após uma `addFileAttachmentAsync`chamada bem-sucedida , `addFileAttachmentFromBase64Async`ou `addItemAttachmentAsync` método, e usá-la em uma chamada de `removeAttachmentAsync` método subsequente. Você também pode chamar `getAttachmentsAsync` (introduzido no conjunto de requisitos 1.8) para obter os anexos e suas IDs para essa sessão de suplemento.

```js
// Removes the specified attachment from the composed item.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the callback function is invoked.
    // Here, the callback function uses an asyncResult parameter and
    // gets the ID of the removed attachment if the removal succeeds.
    // You can optionally pass any object you wish to access in the
    // callback function as an argument to the asyncContext parameter.
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
