---
title: Adicionar e remover os anexos em um suplemento do Outlook
description: Você pode usar várias APIs de anexo para gerenciar os arquivos ou itens do Outlook anexados ao item que o usuário está compondo.
ms.date: 02/24/2021
localization_priority: Normal
ms.openlocfilehash: da426813e865f5607ec3e2c65252e8a406d889e2
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505497"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a><span data-ttu-id="6cb53-103">Gerenciar anexos de um item em um formulário de composição no Outlook</span><span class="sxs-lookup"><span data-stu-id="6cb53-103">Manage an item's attachments in a compose form in Outlook</span></span>

<span data-ttu-id="6cb53-104">A API JavaScript do Office fornece várias APIs que você pode usar para gerenciar anexos de um item quando o usuário está compondo.</span><span class="sxs-lookup"><span data-stu-id="6cb53-104">The Office JavaScript API provides several APIs you can use to manage an item's attachments when the user is composing.</span></span>

## <a name="attach-a-file-or-outlook-item"></a><span data-ttu-id="6cb53-105">Anexar um arquivo ou item do Outlook</span><span class="sxs-lookup"><span data-stu-id="6cb53-105">Attach a file or Outlook item</span></span>

<span data-ttu-id="6cb53-106">Você pode anexar um arquivo ou item do Outlook a um formulário de composição usando o método apropriado para o tipo de anexo.</span><span class="sxs-lookup"><span data-stu-id="6cb53-106">You can attach a file or Outlook item to a compose form by using the method that's appropriate for the type of attachment.</span></span>

- <span data-ttu-id="6cb53-107">[addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): anexar um arquivo</span><span class="sxs-lookup"><span data-stu-id="6cb53-107">[addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file</span></span>
- <span data-ttu-id="6cb53-108">[addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): anexar um arquivo usando sua cadeia de caracteres base64</span><span class="sxs-lookup"><span data-stu-id="6cb53-108">[addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file using its base64 string</span></span>
- <span data-ttu-id="6cb53-109">[addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Anexar um item do Outlook</span><span class="sxs-lookup"><span data-stu-id="6cb53-109">[addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach an Outlook item</span></span>

<span data-ttu-id="6cb53-110">Esses são métodos assíncronos, o que significa que a execução pode continuar sem aguardar a conclusão da ação.</span><span class="sxs-lookup"><span data-stu-id="6cb53-110">These are asynchronous methods, which means execution can go on without waiting for the action to complete.</span></span> <span data-ttu-id="6cb53-111">Dependendo do local original e do tamanho do anexo que está sendo adicionado, a chamada assíncrona pode demorar um pouco para ser concluída.</span><span class="sxs-lookup"><span data-stu-id="6cb53-111">Depending on the original location and size of the attachment being added, the asynchronous call may take a while to complete.</span></span>

<span data-ttu-id="6cb53-112">Se houver tarefas que dependam da conclusão da ação, você deverá executá-las em um método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="6cb53-112">If there are tasks that depend on the action to complete, you should carry out those tasks in a callback method.</span></span> <span data-ttu-id="6cb53-113">Esse método de retorno de chamada é opcional e é invocado quando o carregamento do anexo é concluído.</span><span class="sxs-lookup"><span data-stu-id="6cb53-113">This callback method is optional and is invoked when the attachment upload has completed.</span></span> <span data-ttu-id="6cb53-114">O método de retorno de chamada usa um objeto [AsyncResult](/javascript/api/office/office.asyncresult) como um parâmetro de saída que fornece qualquer status, erro e valor retornado da adição do anexo.</span><span class="sxs-lookup"><span data-stu-id="6cb53-114">The callback method takes an [AsyncResult](/javascript/api/office/office.asyncresult) object as an output parameter that provides any status, error, and returned value from adding the attachment.</span></span> <span data-ttu-id="6cb53-115">Se o retorno de chamada requer parâmetros adicionais, você pode especificá-los no parâmetro opcional `options.asyncContext`.</span><span class="sxs-lookup"><span data-stu-id="6cb53-115">If the callback requires any extra parameters, you can specify them in the optional `options.asyncContext` parameter.</span></span> <span data-ttu-id="6cb53-116">`options.asyncContext` pode ser de qualquer tipo que seu método de retorno de chamada espere.</span><span class="sxs-lookup"><span data-stu-id="6cb53-116">`options.asyncContext` can be of any type that your callback method expects.</span></span>

<span data-ttu-id="6cb53-p103">Por exemplo, você pode definir `options.asyncContext` como um objeto JSON que contém um ou mais pares chave-valor. Você pode encontrar mais exemplos sobre como passar parâmetros opcionais para métodos assíncronos na plataforma de suplementos do Office em [Programação assíncrona em suplementos do Office](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). O exemplo a seguir mostra como usar o parâmetro `asyncContext` para passar dois argumentos para um método de retorno de chamada:</span><span class="sxs-lookup"><span data-stu-id="6cb53-p103">For example, you can define `options.asyncContext` as a JSON object that contains one or more key-value pairs. You can find more examples about passing optional parameters to asynchronous methods in the Office Add-ins platform in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). The following example shows how to use the `asyncContext` parameter to pass 2 arguments to a callback method:</span></span>

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

<span data-ttu-id="6cb53-p104">Você pode verificar o sucesso ou o erro de uma chamada de método assíncrono no método de retorno de chamada usando as propriedades `status` e `error` do objeto `AsyncResult`. Se a anexação for concluída com êxito, você poderá usar a propriedade `AsyncResult.value` para obter a ID do anexo. A ID do anexo é um número inteiro que você pode usar posteriormente para remover o anexo.</span><span class="sxs-lookup"><span data-stu-id="6cb53-p104">You can check for success or error of an asynchronous method call in the callback method using the `status` and `error` properties of the `AsyncResult` object. If the attaching completes successfully, you can use the `AsyncResult.value` property to get the attachment ID. The attachment ID is an integer which you can subsequently use to remove the attachment.</span></span>

> [!NOTE]
> <span data-ttu-id="6cb53-122">A ID do anexo é válida somente dentro da mesma sessão e não tem garantia de mapear para o mesmo anexo entre as sessões.</span><span class="sxs-lookup"><span data-stu-id="6cb53-122">The attachment ID is valid only within the same session and isn't guaranteed to map to the same attachment across sessions.</span></span> <span data-ttu-id="6cb53-123">Os exemplos de quando uma sessão acabou incluem quando o usuário fecha o complemento ou se o usuário começa a compor em um formulário em linha e, subsequentemente, sai do formulário em linha para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="6cb53-123">Examples of when a session is over include when the user closes the add-in, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

### <a name="attach-a-file"></a><span data-ttu-id="6cb53-124">Anexar um arquivo</span><span class="sxs-lookup"><span data-stu-id="6cb53-124">Attach a file</span></span>

<span data-ttu-id="6cb53-125">Você pode anexar um arquivo a uma mensagem ou compromisso em um formulário de composição usando o método e especificando o `addFileAttachmentAsync` URI do arquivo.</span><span class="sxs-lookup"><span data-stu-id="6cb53-125">You can attach a file to a message or appointment in a compose form by using the `addFileAttachmentAsync` method and specifying the URI of the file.</span></span> <span data-ttu-id="6cb53-126">Você também pode usar o `addFileAttachmentFromBase64Async` método, mas especificar a cadeia de caracteres base64 como entrada.</span><span class="sxs-lookup"><span data-stu-id="6cb53-126">You can also use the `addFileAttachmentFromBase64Async` method but specify the base64 string as input.</span></span> <span data-ttu-id="6cb53-127">Se o arquivo estiver protegido, você poderá incluir uma identidade ou um token de autenticação apropriado como um parâmetro de cadeia de caracteres de consulta de URI.</span><span class="sxs-lookup"><span data-stu-id="6cb53-127">If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter.</span></span> <span data-ttu-id="6cb53-128">O Exchange fará uma chamada à URI para obter o anexo, e o serviço Web que protege o arquivo precisará usar o token como um meio de autenticação.</span><span class="sxs-lookup"><span data-stu-id="6cb53-128">Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.</span></span>

<span data-ttu-id="6cb53-p107">O exemplo de JavaScript a seguir é um suplemento de redação que anexa um arquivo, picture.png, de um servidor Web à mensagem ou ao compromisso que está sendo redigido. O método de retorno de chamada usa `asyncResult` como um parâmetro, verifica o status de resultado e obtém a ID do anexo caso o método tenha êxito.</span><span class="sxs-lookup"><span data-stu-id="6cb53-p107">The following JavaScript example is a compose add-in that attaches a file, picture.png, from a web server to the message or appointment being composed. The callback method takes `asyncResult` as a parameter, checks for the result status, and gets the attachment ID if the method succeeds.</span></span>

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

### <a name="attach-an-outlook-item"></a><span data-ttu-id="6cb53-131">Anexar um item do Outlook</span><span class="sxs-lookup"><span data-stu-id="6cb53-131">Attach an Outlook item</span></span>

<span data-ttu-id="6cb53-132">Você pode anexar um item do Outlook (por exemplo, email, calendário ou item de contato) a uma mensagem ou compromisso em um formulário de redação especificando a ID dos Serviços Web do Exchange (EWS) do item e usando o `addItemAttachmentAsync` método.</span><span class="sxs-lookup"><span data-stu-id="6cb53-132">You can attach an Outlook item (for example, email, calendar, or contact item) to a message or appointment in a compose form by specifying the Exchange Web Services (EWS) ID of the item and using the `addItemAttachmentAsync` method.</span></span> <span data-ttu-id="6cb53-133">Você pode obter a ID do EWS de um email, calendário, contato ou item de tarefa na caixa de correio do usuário usando o método [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) e acessando a operação EWS [FindItem](/exchange/client-developer/web-service-reference/finditem-operation).</span><span class="sxs-lookup"><span data-stu-id="6cb53-133">You can get the EWS ID of an email, calendar, contact, or task item in the user's mailbox by using the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method and accessing the EWS operation [FindItem](/exchange/client-developer/web-service-reference/finditem-operation).</span></span> <span data-ttu-id="6cb53-134">A propriedade [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) também fornece a ID dos EWS de um item existente em um formulário de leitura.</span><span class="sxs-lookup"><span data-stu-id="6cb53-134">The [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property also provides the EWS ID of an existing item in a read form.</span></span>

<span data-ttu-id="6cb53-135">A função JavaScript a seguir, , estende o primeiro exemplo acima e adiciona um item como um anexo ao email ou compromisso `addItemAttachment` que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="6cb53-135">The following JavaScript function, `addItemAttachment`, extends the first example above, and adds an item as an attachment to the email or appointment that is being composed.</span></span> <span data-ttu-id="6cb53-136">A função utiliza como argumento a ID dos EWS do item que será anexado.</span><span class="sxs-lookup"><span data-stu-id="6cb53-136">The function takes as an argument the EWS ID of the item that is to be attached.</span></span> <span data-ttu-id="6cb53-137">Se a anexação for bem-sucedida, ela obtém a ID do anexo para processamento posterior, incluindo a remoção desse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="6cb53-137">If attaching succeeds, it gets the attachment ID for further processing, including removing that attachment in the same session.</span></span>

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
> <span data-ttu-id="6cb53-138">Você pode usar um complemento de composição para anexar uma instância de um compromisso recorrente no Outlook na Web ou em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="6cb53-138">You can use a compose add-in to attach an instance of a recurring appointment in Outlook on the web or on mobile devices.</span></span> <span data-ttu-id="6cb53-139">No entanto, em um cliente de área de trabalho do Outlook com suporte, tentar anexar uma instância resultaria em anexar a série recorrente (o compromisso pai).</span><span class="sxs-lookup"><span data-stu-id="6cb53-139">However, in a supporting Outlook desktop client, attempting to attach an instance would result in attaching the recurring series (the parent appointment).</span></span>

## <a name="get-attachments"></a><span data-ttu-id="6cb53-140">Obter anexos</span><span class="sxs-lookup"><span data-stu-id="6cb53-140">Get attachments</span></span>

<span data-ttu-id="6cb53-141">APIs para obter anexos no modo de redação estão disponíveis no conjunto [de requisitos 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="6cb53-141">APIs to get attachments in compose mode are available from [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

- [<span data-ttu-id="6cb53-142">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="6cb53-142">getAttachmentsAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="6cb53-143">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="6cb53-143">getAttachmentContentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

<span data-ttu-id="6cb53-144">Você pode usar o [método getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) para obter os anexos da mensagem ou do compromisso que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="6cb53-144">You can use the [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method to get the attachments of the message or appointment being composed.</span></span>

<span data-ttu-id="6cb53-145">Para obter o conteúdo de um anexo, você pode usar o [método getAttachmentContentAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)</span><span class="sxs-lookup"><span data-stu-id="6cb53-145">To get an attachment's content, you can use the [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="6cb53-146">Os formatos com suporte estão listados na [enumeração AttachmentContentFormat.](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)</span><span class="sxs-lookup"><span data-stu-id="6cb53-146">The supported formats are listed in the [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) enum.</span></span>

<span data-ttu-id="6cb53-147">Você deve fornecer um método de retorno de chamada para verificar o status e qualquer erro usando o `AsyncResult` objeto de parâmetro de saída.</span><span class="sxs-lookup"><span data-stu-id="6cb53-147">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="6cb53-148">Você também pode passar quaisquer parâmetros adicionais para o método de retorno de chamada usando o parâmetro `asyncContext` opcional.</span><span class="sxs-lookup"><span data-stu-id="6cb53-148">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="6cb53-149">O exemplo JavaScript a seguir obtém os anexos e permite configurar o tratamento distinto para cada formato de anexo com suporte.</span><span class="sxs-lookup"><span data-stu-id="6cb53-149">The following JavaScript example gets the attachments and allows you to set up distinct handling for each supported attachment format.</span></span>

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

## <a name="remove-an-attachment"></a><span data-ttu-id="6cb53-150">Remover um anexo</span><span class="sxs-lookup"><span data-stu-id="6cb53-150">Remove an attachment</span></span>

<span data-ttu-id="6cb53-151">Você pode remover um anexo de arquivo ou item de uma mensagem ou item de compromisso em um formulário de composição especificando a ID de anexo correspondente ao usar o [método removeAttachmentAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)</span><span class="sxs-lookup"><span data-stu-id="6cb53-151">You can remove a file or item attachment from a message or appointment item in a compose form by specifying the corresponding attachment ID when using the [removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6cb53-152">Se você estiver usando o conjunto de requisitos 1.7 ou anterior, você deve remover apenas anexos que o mesmo complemento adicionou na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="6cb53-152">If you're using requirement set 1.7 or earlier, you should only remove attachments that the same add-in has added in the same session.</span></span>

<span data-ttu-id="6cb53-153">Semelhante aos `addFileAttachmentAsync` métodos , e é um método `addItemAttachmentAsync` `getAttachmentsAsync` `removeAttachmentAsync` assíncrono.</span><span class="sxs-lookup"><span data-stu-id="6cb53-153">Similar to the `addFileAttachmentAsync`, `addItemAttachmentAsync`, and `getAttachmentsAsync` methods, `removeAttachmentAsync` is an asynchronous method.</span></span> <span data-ttu-id="6cb53-154">Você deve fornecer um método de retorno de chamada para verificar o status e qualquer erro usando o `AsyncResult` objeto de parâmetro de saída.</span><span class="sxs-lookup"><span data-stu-id="6cb53-154">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="6cb53-155">Você também pode passar quaisquer parâmetros adicionais para o método de retorno de chamada usando o parâmetro `asyncContext` opcional.</span><span class="sxs-lookup"><span data-stu-id="6cb53-155">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="6cb53-156">A função JavaScript a seguir, , continua a estender os exemplos acima e remove o anexo especificado do email ou compromisso `removeAttachment` que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="6cb53-156">The following JavaScript function, `removeAttachment`, continues to extend the examples above, and removes the specified attachment from the email or appointment that is being composed.</span></span> <span data-ttu-id="6cb53-157">A função utiliza como argumento a ID do anexo a ser removido.</span><span class="sxs-lookup"><span data-stu-id="6cb53-157">The function takes as an argument the ID of the attachment to be removed.</span></span> <span data-ttu-id="6cb53-158">Você pode obter a ID de um anexo após uma chamada bem-sucedida , ou método, e `addFileAttachmentAsync` usá-lo em uma chamada de método `addFileAttachmentFromBase64Async` `addItemAttachmentAsync` `removeAttachmentAsync` subsequente.</span><span class="sxs-lookup"><span data-stu-id="6cb53-158">You can obtain the ID of an attachment after a successful `addFileAttachmentAsync`, `addFileAttachmentFromBase64Async`, or `addItemAttachmentAsync` method call, and use it in a subsequent `removeAttachmentAsync` method call.</span></span> <span data-ttu-id="6cb53-159">Você também pode chamar (introduzido no conjunto de requisitos 1.8) para obter os anexos e suas IDs para essa sessão `getAttachmentsAsync` de complemento.</span><span class="sxs-lookup"><span data-stu-id="6cb53-159">You can also call `getAttachmentsAsync` (introduced in requirement set 1.8) to get the attachments and their IDs for that add-in session.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="6cb53-160">Confira também</span><span class="sxs-lookup"><span data-stu-id="6cb53-160">See also</span></span>

- [<span data-ttu-id="6cb53-161">Criar suplementos do Outlook para formulários de redação</span><span class="sxs-lookup"><span data-stu-id="6cb53-161">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="6cb53-162">Programação assíncrona em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6cb53-162">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
