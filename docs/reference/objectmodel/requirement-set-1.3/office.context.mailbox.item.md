---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,3
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 6896c849b144f3720a3d9fb284e88d18d5c8ff43
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/14/2019
ms.locfileid: "30600295"
---
# <a name="item"></a><span data-ttu-id="20889-102">item</span><span class="sxs-lookup"><span data-stu-id="20889-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="20889-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="20889-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="20889-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="20889-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-106">Requirements</span></span>

|<span data-ttu-id="20889-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-107">Requirement</span></span>| <span data-ttu-id="20889-108">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-110">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-110">1.0</span></span>|
|[<span data-ttu-id="20889-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="20889-112">Restricted</span></span>|
|[<span data-ttu-id="20889-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-114">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="20889-115">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-115">Example</span></span>

<span data-ttu-id="20889-116">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="20889-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```

### <a name="members"></a><span data-ttu-id="20889-117">Membros</span><span class="sxs-lookup"><span data-stu-id="20889-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="20889-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="20889-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="20889-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="20889-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-121">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="20889-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="20889-122">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="20889-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="20889-123">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-123">Type</span></span>

*   <span data-ttu-id="20889-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="20889-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-125">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-125">Requirements</span></span>

|<span data-ttu-id="20889-126">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-126">Requirement</span></span>| <span data-ttu-id="20889-127">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-128">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-129">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-129">1.0</span></span>|
|[<span data-ttu-id="20889-130">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-131">ReadItem</span></span>|
|[<span data-ttu-id="20889-132">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="20889-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-133">Read</span><span class="sxs-lookup"><span data-stu-id="20889-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-134">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-134">Example</span></span>

<span data-ttu-id="20889-135">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="20889-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="20889-136">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="20889-136">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="20889-137">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="20889-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="20889-138">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="20889-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-139">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-139">Type</span></span>

*   [<span data-ttu-id="20889-140">Destinatários</span><span class="sxs-lookup"><span data-stu-id="20889-140">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="20889-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-141">Requirements</span></span>

|<span data-ttu-id="20889-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-142">Requirement</span></span>| <span data-ttu-id="20889-143">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-145">1.1</span><span class="sxs-lookup"><span data-stu-id="20889-145">1.1</span></span>|
|[<span data-ttu-id="20889-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-147">ReadItem</span></span>|
|[<span data-ttu-id="20889-148">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="20889-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-149">Escrever</span><span class="sxs-lookup"><span data-stu-id="20889-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-150">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="20889-151">body :[Body](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="20889-151">body :[Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="20889-152">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="20889-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-153">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-153">Type</span></span>

*   [<span data-ttu-id="20889-154">Body</span><span class="sxs-lookup"><span data-stu-id="20889-154">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="20889-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-155">Requirements</span></span>

|<span data-ttu-id="20889-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-156">Requirement</span></span>| <span data-ttu-id="20889-157">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-159">1.1</span><span class="sxs-lookup"><span data-stu-id="20889-159">1.1</span></span>|
|[<span data-ttu-id="20889-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-161">ReadItem</span></span>|
|[<span data-ttu-id="20889-162">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="20889-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-163">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-164">Example</span></span>

<span data-ttu-id="20889-165">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="20889-165">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="20889-166">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="20889-166">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="20889-167">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="20889-167">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="20889-168">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="20889-168">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="20889-169">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="20889-169">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="20889-170">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="20889-170">Read mode</span></span>

<span data-ttu-id="20889-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="20889-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="20889-173">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="20889-173">Compose mode</span></span>

<span data-ttu-id="20889-174">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="20889-174">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="20889-175">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-175">Type</span></span>

*   <span data-ttu-id="20889-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="20889-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-177">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-177">Requirements</span></span>

|<span data-ttu-id="20889-178">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-178">Requirement</span></span>| <span data-ttu-id="20889-179">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-180">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-181">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-181">1.0</span></span>|
|[<span data-ttu-id="20889-182">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-182">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-183">ReadItem</span></span>|
|[<span data-ttu-id="20889-184">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="20889-184">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-185">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-185">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="20889-186">(anulável) conversationId :Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="20889-186">(nullable) conversationId :String</span></span>

<span data-ttu-id="20889-187">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="20889-187">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="20889-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="20889-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="20889-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="20889-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-192">Type</span></span>

*   <span data-ttu-id="20889-193">String</span><span class="sxs-lookup"><span data-stu-id="20889-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-194">Requirements</span></span>

|<span data-ttu-id="20889-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-195">Requirement</span></span>| <span data-ttu-id="20889-196">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-198">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-198">1.0</span></span>|
|[<span data-ttu-id="20889-199">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-199">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-200">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-200">ReadItem</span></span>|
|[<span data-ttu-id="20889-201">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="20889-201">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-202">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-203">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-203">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="20889-204">dateTimeCreated :Data</span><span class="sxs-lookup"><span data-stu-id="20889-204">dateTimeCreated :Date</span></span>

<span data-ttu-id="20889-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="20889-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-207">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-207">Type</span></span>

*   <span data-ttu-id="20889-208">Data</span><span class="sxs-lookup"><span data-stu-id="20889-208">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-209">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-209">Requirements</span></span>

|<span data-ttu-id="20889-210">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-210">Requirement</span></span>| <span data-ttu-id="20889-211">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-212">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-213">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-213">1.0</span></span>|
|[<span data-ttu-id="20889-214">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-214">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-215">ReadItem</span></span>|
|[<span data-ttu-id="20889-216">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-217">Read</span><span class="sxs-lookup"><span data-stu-id="20889-217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-218">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-218">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="20889-219">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="20889-219">dateTimeModified :Date</span></span>

<span data-ttu-id="20889-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="20889-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-222">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="20889-222">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-223">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-223">Type</span></span>

*   <span data-ttu-id="20889-224">Data</span><span class="sxs-lookup"><span data-stu-id="20889-224">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-225">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-225">Requirements</span></span>

|<span data-ttu-id="20889-226">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-226">Requirement</span></span>| <span data-ttu-id="20889-227">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-228">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-229">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-229">1.0</span></span>|
|[<span data-ttu-id="20889-230">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-230">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-231">ReadItem</span></span>|
|[<span data-ttu-id="20889-232">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-232">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-233">Read</span><span class="sxs-lookup"><span data-stu-id="20889-233">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-234">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-234">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="20889-235">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="20889-235">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="20889-236">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="20889-236">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="20889-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="20889-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="20889-239">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="20889-239">Read mode</span></span>

<span data-ttu-id="20889-240">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="20889-240">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="20889-241">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="20889-241">Compose mode</span></span>

<span data-ttu-id="20889-242">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="20889-242">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="20889-243">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="20889-243">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="20889-244">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="20889-244">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="20889-245">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-245">Type</span></span>

*   <span data-ttu-id="20889-246">Data | [Hora](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="20889-246">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-247">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-247">Requirements</span></span>

|<span data-ttu-id="20889-248">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-248">Requirement</span></span>| <span data-ttu-id="20889-249">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-250">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-251">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-251">1.0</span></span>|
|[<span data-ttu-id="20889-252">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-252">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-253">ReadItem</span></span>|
|[<span data-ttu-id="20889-254">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-254">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-255">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-255">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="20889-256">De:[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="20889-256">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="20889-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="20889-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="20889-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="20889-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-261">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="20889-261">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-262">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-262">Type</span></span>

*   [<span data-ttu-id="20889-263">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="20889-263">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="20889-264">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-264">Requirements</span></span>

|<span data-ttu-id="20889-265">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-265">Requirement</span></span>| <span data-ttu-id="20889-266">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-267">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-268">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-268">1.0</span></span>|
|[<span data-ttu-id="20889-269">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-270">ReadItem</span></span>|
|[<span data-ttu-id="20889-271">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-272">Read</span><span class="sxs-lookup"><span data-stu-id="20889-272">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-273">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-273">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="20889-274">internetMessageId Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="20889-274">internetMessageId :String</span></span>

<span data-ttu-id="20889-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="20889-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-277">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-277">Type</span></span>

*   <span data-ttu-id="20889-278">String</span><span class="sxs-lookup"><span data-stu-id="20889-278">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-279">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-279">Requirements</span></span>

|<span data-ttu-id="20889-280">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-280">Requirement</span></span>| <span data-ttu-id="20889-281">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-282">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-283">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-283">1.0</span></span>|
|[<span data-ttu-id="20889-284">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-284">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-285">ReadItem</span></span>|
|[<span data-ttu-id="20889-286">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-286">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-287">Read</span><span class="sxs-lookup"><span data-stu-id="20889-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-288">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-288">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="20889-289">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="20889-289">itemClass :String</span></span>

<span data-ttu-id="20889-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="20889-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="20889-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="20889-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="20889-294">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-294">Type</span></span> | <span data-ttu-id="20889-295">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-295">Description</span></span> | <span data-ttu-id="20889-296">classe de item</span><span class="sxs-lookup"><span data-stu-id="20889-296">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="20889-297">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="20889-297">Appointment items</span></span> | <span data-ttu-id="20889-298">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="20889-298">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="20889-299">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="20889-299">Message items</span></span> | <span data-ttu-id="20889-300">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="20889-300">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="20889-301">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="20889-301">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-302">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-302">Type</span></span>

*   <span data-ttu-id="20889-303">String</span><span class="sxs-lookup"><span data-stu-id="20889-303">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-304">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-304">Requirements</span></span>

|<span data-ttu-id="20889-305">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-305">Requirement</span></span>| <span data-ttu-id="20889-306">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-307">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-308">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-308">1.0</span></span>|
|[<span data-ttu-id="20889-309">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-310">ReadItem</span></span>|
|[<span data-ttu-id="20889-311">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-312">Read</span><span class="sxs-lookup"><span data-stu-id="20889-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-313">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-313">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="20889-314">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="20889-314">(nullable) itemId :String</span></span>

<span data-ttu-id="20889-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="20889-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-317">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="20889-317">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="20889-318">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="20889-318">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="20889-319">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="20889-319">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="20889-320">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="20889-320">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="20889-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="20889-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-323">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-323">Type</span></span>

*   <span data-ttu-id="20889-324">String</span><span class="sxs-lookup"><span data-stu-id="20889-324">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-325">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-325">Requirements</span></span>

|<span data-ttu-id="20889-326">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-326">Requirement</span></span>| <span data-ttu-id="20889-327">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-328">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-329">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-329">1.0</span></span>|
|[<span data-ttu-id="20889-330">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-331">ReadItem</span></span>|
|[<span data-ttu-id="20889-332">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-333">Read</span><span class="sxs-lookup"><span data-stu-id="20889-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-334">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-334">Example</span></span>

<span data-ttu-id="20889-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="20889-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="20889-337">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="20889-337">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="20889-338">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="20889-338">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="20889-339">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="20889-339">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-340">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-340">Type</span></span>

*   [<span data-ttu-id="20889-341">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="20889-341">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="20889-342">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-342">Requirements</span></span>

|<span data-ttu-id="20889-343">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-343">Requirement</span></span>| <span data-ttu-id="20889-344">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-344">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-345">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-345">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-346">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-346">1.0</span></span>|
|[<span data-ttu-id="20889-347">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-347">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-348">ReadItem</span></span>|
|[<span data-ttu-id="20889-349">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-349">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-350">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-350">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-351">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-351">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="20889-352">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="20889-352">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="20889-353">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="20889-353">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="20889-354">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="20889-354">Read mode</span></span>

<span data-ttu-id="20889-355">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="20889-355">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="20889-356">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="20889-356">Compose mode</span></span>

<span data-ttu-id="20889-357">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="20889-357">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="20889-358">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-358">Type</span></span>

*   <span data-ttu-id="20889-359">Cadeia de caracteres | [Localização](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="20889-359">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-360">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-360">Requirements</span></span>

|<span data-ttu-id="20889-361">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-361">Requirement</span></span>| <span data-ttu-id="20889-362">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-363">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-364">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-364">1.0</span></span>|
|[<span data-ttu-id="20889-365">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-365">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-366">ReadItem</span></span>|
|[<span data-ttu-id="20889-367">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-367">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-368">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-368">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="20889-369">normalizedSubject :Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="20889-369">normalizedSubject :String</span></span>

<span data-ttu-id="20889-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="20889-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="20889-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="20889-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-374">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-374">Type</span></span>

*   <span data-ttu-id="20889-375">String</span><span class="sxs-lookup"><span data-stu-id="20889-375">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-376">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-376">Requirements</span></span>

|<span data-ttu-id="20889-377">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-377">Requirement</span></span>| <span data-ttu-id="20889-378">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-378">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-379">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-380">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-380">1.0</span></span>|
|[<span data-ttu-id="20889-381">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-381">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-382">ReadItem</span></span>|
|[<span data-ttu-id="20889-383">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-383">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-384">Read</span><span class="sxs-lookup"><span data-stu-id="20889-384">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-385">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-385">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="20889-386">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="20889-386">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="20889-387">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="20889-387">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-388">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-388">Type</span></span>

*   [<span data-ttu-id="20889-389">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="20889-389">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="20889-390">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-390">Requirements</span></span>

|<span data-ttu-id="20889-391">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-391">Requirement</span></span>| <span data-ttu-id="20889-392">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-392">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-393">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-394">1.3</span><span class="sxs-lookup"><span data-stu-id="20889-394">1.3</span></span>|
|[<span data-ttu-id="20889-395">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-395">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-396">ReadItem</span></span>|
|[<span data-ttu-id="20889-397">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-397">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-398">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-398">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-399">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-399">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="20889-400">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="20889-400">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="20889-401">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="20889-401">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="20889-402">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="20889-402">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="20889-403">Modo de Leitura</span><span class="sxs-lookup"><span data-stu-id="20889-403">Read mode</span></span>

<span data-ttu-id="20889-404">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="20889-404">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="20889-405">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="20889-405">Compose mode</span></span>

<span data-ttu-id="20889-406">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="20889-406">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="20889-407">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-407">Type</span></span>

*   <span data-ttu-id="20889-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="20889-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-409">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-409">Requirements</span></span>

|<span data-ttu-id="20889-410">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-410">Requirement</span></span>| <span data-ttu-id="20889-411">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-412">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-413">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-413">1.0</span></span>|
|[<span data-ttu-id="20889-414">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-414">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-415">ReadItem</span></span>|
|[<span data-ttu-id="20889-416">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-416">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-417">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-417">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="20889-418">organizador:[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="20889-418">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="20889-p124">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="20889-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-421">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-421">Type</span></span>

*   [<span data-ttu-id="20889-422">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="20889-422">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="20889-423">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-423">Requirements</span></span>

|<span data-ttu-id="20889-424">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-424">Requirement</span></span>| <span data-ttu-id="20889-425">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-426">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-427">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-427">1.0</span></span>|
|[<span data-ttu-id="20889-428">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-428">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-429">ReadItem</span></span>|
|[<span data-ttu-id="20889-430">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-430">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-431">Read</span><span class="sxs-lookup"><span data-stu-id="20889-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-432">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-432">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="20889-433">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="20889-433">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="20889-434">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="20889-434">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="20889-435">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="20889-435">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="20889-436">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="20889-436">Read mode</span></span>

<span data-ttu-id="20889-437">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="20889-437">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="20889-438">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="20889-438">Compose mode</span></span>

<span data-ttu-id="20889-439">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="20889-439">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="20889-440">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-440">Type</span></span>

*   <span data-ttu-id="20889-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="20889-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-442">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-442">Requirements</span></span>

|<span data-ttu-id="20889-443">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-443">Requirement</span></span>| <span data-ttu-id="20889-444">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-445">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-446">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-446">1.0</span></span>|
|[<span data-ttu-id="20889-447">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-447">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-448">ReadItem</span></span>|
|[<span data-ttu-id="20889-449">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-449">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-450">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-450">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="20889-451">remetente :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="20889-451">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="20889-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="20889-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="20889-p127">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="20889-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-456">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="20889-456">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="20889-457">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-457">Type</span></span>

*   [<span data-ttu-id="20889-458">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="20889-458">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="20889-459">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-459">Requirements</span></span>

|<span data-ttu-id="20889-460">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-460">Requirement</span></span>| <span data-ttu-id="20889-461">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-462">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-463">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-463">1.0</span></span>|
|[<span data-ttu-id="20889-464">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-464">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-465">ReadItem</span></span>|
|[<span data-ttu-id="20889-466">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-466">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-467">Read</span><span class="sxs-lookup"><span data-stu-id="20889-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-468">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-468">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="20889-469">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="20889-469">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="20889-470">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="20889-470">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="20889-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="20889-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="20889-473">Modo de Leitura</span><span class="sxs-lookup"><span data-stu-id="20889-473">Read mode</span></span>

<span data-ttu-id="20889-474">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="20889-474">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="20889-475">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="20889-475">Compose mode</span></span>

<span data-ttu-id="20889-476">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="20889-476">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="20889-477">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="20889-477">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="20889-478">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="20889-478">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="20889-479">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-479">Type</span></span>

*   <span data-ttu-id="20889-480">Data | [Hora](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="20889-480">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-481">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-481">Requirements</span></span>

|<span data-ttu-id="20889-482">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-482">Requirement</span></span>| <span data-ttu-id="20889-483">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-484">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-485">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-485">1.0</span></span>|
|[<span data-ttu-id="20889-486">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-486">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-487">ReadItem</span></span>|
|[<span data-ttu-id="20889-488">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-488">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-489">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-489">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="20889-490">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="20889-490">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="20889-491">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="20889-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="20889-492">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="20889-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="20889-493">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="20889-493">Read mode</span></span>

<span data-ttu-id="20889-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="20889-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="20889-496">Modo de composição</span><span class="sxs-lookup"><span data-stu-id="20889-496">Compose mode</span></span>

<span data-ttu-id="20889-497">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="20889-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="20889-498">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-498">Type</span></span>

*   <span data-ttu-id="20889-499">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="20889-499">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-500">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-500">Requirements</span></span>

|<span data-ttu-id="20889-501">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-501">Requirement</span></span>| <span data-ttu-id="20889-502">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-503">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-504">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-504">1.0</span></span>|
|[<span data-ttu-id="20889-505">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-506">ReadItem</span></span>|
|[<span data-ttu-id="20889-507">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-508">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-508">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="20889-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="20889-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="20889-510">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="20889-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="20889-511">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="20889-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="20889-512">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="20889-512">Read mode</span></span>

<span data-ttu-id="20889-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="20889-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="20889-515">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="20889-515">Compose mode</span></span>

<span data-ttu-id="20889-516">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="20889-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="20889-517">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-517">Type</span></span>

*   <span data-ttu-id="20889-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="20889-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-519">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-519">Requirements</span></span>

|<span data-ttu-id="20889-520">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-520">Requirement</span></span>| <span data-ttu-id="20889-521">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-522">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-523">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-523">1.0</span></span>|
|[<span data-ttu-id="20889-524">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-525">ReadItem</span></span>|
|[<span data-ttu-id="20889-526">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-526">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-527">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-527">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="20889-528">Métodos</span><span class="sxs-lookup"><span data-stu-id="20889-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="20889-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="20889-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="20889-530">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="20889-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="20889-531">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="20889-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="20889-532">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="20889-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="20889-533">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="20889-533">Parameters</span></span>

|<span data-ttu-id="20889-534">Name</span><span class="sxs-lookup"><span data-stu-id="20889-534">Name</span></span>| <span data-ttu-id="20889-535">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-535">Type</span></span>| <span data-ttu-id="20889-536">Atributos</span><span class="sxs-lookup"><span data-stu-id="20889-536">Attributes</span></span>| <span data-ttu-id="20889-537">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="20889-538">String</span><span class="sxs-lookup"><span data-stu-id="20889-538">String</span></span>||<span data-ttu-id="20889-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="20889-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="20889-541">String</span><span class="sxs-lookup"><span data-stu-id="20889-541">String</span></span>||<span data-ttu-id="20889-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="20889-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="20889-544">Object</span><span class="sxs-lookup"><span data-stu-id="20889-544">Object</span></span>| <span data-ttu-id="20889-545">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-545">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-546">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="20889-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="20889-547">Objeto</span><span class="sxs-lookup"><span data-stu-id="20889-547">Object</span></span>| <span data-ttu-id="20889-548">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-548">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-549">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="20889-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="20889-550">function</span><span class="sxs-lookup"><span data-stu-id="20889-550">function</span></span>| <span data-ttu-id="20889-551">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-551">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-552">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="20889-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="20889-553">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="20889-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="20889-554">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="20889-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="20889-555">Erros</span><span class="sxs-lookup"><span data-stu-id="20889-555">Errors</span></span>

| <span data-ttu-id="20889-556">Código de erro</span><span class="sxs-lookup"><span data-stu-id="20889-556">Error code</span></span> | <span data-ttu-id="20889-557">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="20889-558">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="20889-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="20889-559">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="20889-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="20889-560">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="20889-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="20889-561">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-561">Requirements</span></span>

|<span data-ttu-id="20889-562">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-562">Requirement</span></span>| <span data-ttu-id="20889-563">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-564">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-565">1.1</span><span class="sxs-lookup"><span data-stu-id="20889-565">1.1</span></span>|
|[<span data-ttu-id="20889-566">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="20889-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="20889-568">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-569">Escrever</span><span class="sxs-lookup"><span data-stu-id="20889-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-570">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-570">Example</span></span>

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="20889-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="20889-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="20889-572">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="20889-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="20889-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="20889-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="20889-576">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="20889-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="20889-577">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="20889-577">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="20889-578">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="20889-578">Parameters</span></span>

|<span data-ttu-id="20889-579">Name</span><span class="sxs-lookup"><span data-stu-id="20889-579">Name</span></span>| <span data-ttu-id="20889-580">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-580">Type</span></span>| <span data-ttu-id="20889-581">Atributos</span><span class="sxs-lookup"><span data-stu-id="20889-581">Attributes</span></span>| <span data-ttu-id="20889-582">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="20889-583">String</span><span class="sxs-lookup"><span data-stu-id="20889-583">String</span></span>||<span data-ttu-id="20889-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="20889-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="20889-586">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="20889-586">String</span></span>||<span data-ttu-id="20889-587">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="20889-587">The subject of the item to be attached.</span></span> <span data-ttu-id="20889-588">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="20889-588">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="20889-589">Object</span><span class="sxs-lookup"><span data-stu-id="20889-589">Object</span></span>| <span data-ttu-id="20889-590">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-590">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-591">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="20889-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="20889-592">Object</span><span class="sxs-lookup"><span data-stu-id="20889-592">Object</span></span>| <span data-ttu-id="20889-593">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-593">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-594">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="20889-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="20889-595">function</span><span class="sxs-lookup"><span data-stu-id="20889-595">function</span></span>| <span data-ttu-id="20889-596">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-596">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-597">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="20889-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="20889-598">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="20889-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="20889-599">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="20889-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="20889-600">Erros</span><span class="sxs-lookup"><span data-stu-id="20889-600">Errors</span></span>

| <span data-ttu-id="20889-601">Código de erro</span><span class="sxs-lookup"><span data-stu-id="20889-601">Error code</span></span> | <span data-ttu-id="20889-602">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="20889-603">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="20889-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="20889-604">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-604">Requirements</span></span>

|<span data-ttu-id="20889-605">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-605">Requirement</span></span>| <span data-ttu-id="20889-606">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-607">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-608">1.1</span><span class="sxs-lookup"><span data-stu-id="20889-608">1.1</span></span>|
|[<span data-ttu-id="20889-609">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="20889-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="20889-611">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-612">Escrever</span><span class="sxs-lookup"><span data-stu-id="20889-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-613">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-613">Example</span></span>

<span data-ttu-id="20889-614">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="20889-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="20889-615">close()</span><span class="sxs-lookup"><span data-stu-id="20889-615">close()</span></span>

<span data-ttu-id="20889-616">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="20889-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="20889-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="20889-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-619">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="20889-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="20889-620">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="20889-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-621">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-621">Requirements</span></span>

|<span data-ttu-id="20889-622">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-622">Requirement</span></span>| <span data-ttu-id="20889-623">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-624">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-625">1.3</span><span class="sxs-lookup"><span data-stu-id="20889-625">1.3</span></span>|
|[<span data-ttu-id="20889-626">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-627">Restrito</span><span class="sxs-lookup"><span data-stu-id="20889-627">Restricted</span></span>|
|[<span data-ttu-id="20889-628">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-629">Escrever</span><span class="sxs-lookup"><span data-stu-id="20889-629">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="20889-630">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="20889-630">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="20889-631">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="20889-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-632">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="20889-632">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="20889-633">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="20889-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="20889-634">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="20889-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="20889-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="20889-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="20889-638">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="20889-638">Parameters</span></span>

|<span data-ttu-id="20889-639">Name</span><span class="sxs-lookup"><span data-stu-id="20889-639">Name</span></span>| <span data-ttu-id="20889-640">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-640">Type</span></span>| <span data-ttu-id="20889-641">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="20889-642">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="20889-642">String &#124; Object</span></span>| |<span data-ttu-id="20889-643">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="20889-643">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="20889-644">A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="20889-644">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="20889-645">**OU**</span><span class="sxs-lookup"><span data-stu-id="20889-645">**OR**</span></span><br/><span data-ttu-id="20889-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="20889-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="20889-648">String</span><span class="sxs-lookup"><span data-stu-id="20889-648">String</span></span> | <span data-ttu-id="20889-649">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-649">&lt;optional&gt;</span></span> | <span data-ttu-id="20889-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="20889-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="20889-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="20889-653">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-653">&lt;optional&gt;</span></span> | <span data-ttu-id="20889-654">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="20889-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="20889-655">String</span><span class="sxs-lookup"><span data-stu-id="20889-655">String</span></span> | | <span data-ttu-id="20889-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="20889-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="20889-658">String</span><span class="sxs-lookup"><span data-stu-id="20889-658">String</span></span> | | <span data-ttu-id="20889-659">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="20889-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="20889-660">String</span><span class="sxs-lookup"><span data-stu-id="20889-660">String</span></span> | | <span data-ttu-id="20889-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="20889-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="20889-663">String</span><span class="sxs-lookup"><span data-stu-id="20889-663">String</span></span> | | <span data-ttu-id="20889-p144">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="20889-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="20889-667">function</span><span class="sxs-lookup"><span data-stu-id="20889-667">function</span></span> | <span data-ttu-id="20889-668">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-668">&lt;optional&gt;</span></span> | <span data-ttu-id="20889-669">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="20889-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="20889-670">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-670">Requirements</span></span>

|<span data-ttu-id="20889-671">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-671">Requirement</span></span>| <span data-ttu-id="20889-672">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-673">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-674">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-674">1.0</span></span>|
|[<span data-ttu-id="20889-675">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-676">ReadItem</span></span>|
|[<span data-ttu-id="20889-677">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="20889-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-678">Read</span><span class="sxs-lookup"><span data-stu-id="20889-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="20889-679">Exemplos</span><span class="sxs-lookup"><span data-stu-id="20889-679">Examples</span></span>

<span data-ttu-id="20889-680">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="20889-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="20889-681">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="20889-681">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="20889-682">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="20889-682">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="20889-683">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="20889-683">Reply with a body and a file attachment.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="20889-684">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="20889-684">Reply with a body and an item attachment.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="20889-685">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="20889-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="20889-686">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="20889-686">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="20889-687">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="20889-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-688">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="20889-688">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="20889-689">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="20889-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="20889-690">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="20889-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="20889-p145">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="20889-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="20889-694">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="20889-694">Parameters</span></span>

|<span data-ttu-id="20889-695">Name</span><span class="sxs-lookup"><span data-stu-id="20889-695">Name</span></span>| <span data-ttu-id="20889-696">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-696">Type</span></span>| <span data-ttu-id="20889-697">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="20889-698">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="20889-698">String &#124; Object</span></span>| | <span data-ttu-id="20889-699">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="20889-699">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="20889-700">A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="20889-700">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="20889-701">**OU**</span><span class="sxs-lookup"><span data-stu-id="20889-701">**OR**</span></span><br/><span data-ttu-id="20889-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="20889-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="20889-704">String</span><span class="sxs-lookup"><span data-stu-id="20889-704">String</span></span> | <span data-ttu-id="20889-705">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-705">&lt;optional&gt;</span></span> | <span data-ttu-id="20889-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="20889-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="20889-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="20889-709">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-709">&lt;optional&gt;</span></span> | <span data-ttu-id="20889-710">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="20889-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="20889-711">String</span><span class="sxs-lookup"><span data-stu-id="20889-711">String</span></span> | | <span data-ttu-id="20889-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="20889-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="20889-714">String</span><span class="sxs-lookup"><span data-stu-id="20889-714">String</span></span> | | <span data-ttu-id="20889-715">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="20889-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="20889-716">String</span><span class="sxs-lookup"><span data-stu-id="20889-716">String</span></span> | | <span data-ttu-id="20889-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="20889-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="20889-719">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="20889-719">String</span></span> | | <span data-ttu-id="20889-p151">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="20889-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="20889-723">function</span><span class="sxs-lookup"><span data-stu-id="20889-723">function</span></span> | <span data-ttu-id="20889-724">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-724">&lt;optional&gt;</span></span> | <span data-ttu-id="20889-725">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="20889-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="20889-726">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-726">Requirements</span></span>

|<span data-ttu-id="20889-727">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-727">Requirement</span></span>| <span data-ttu-id="20889-728">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-729">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-730">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-730">1.0</span></span>|
|[<span data-ttu-id="20889-731">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-732">ReadItem</span></span>|
|[<span data-ttu-id="20889-733">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-734">Read</span><span class="sxs-lookup"><span data-stu-id="20889-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="20889-735">Exemplos</span><span class="sxs-lookup"><span data-stu-id="20889-735">Examples</span></span>

<span data-ttu-id="20889-736">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="20889-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="20889-737">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="20889-737">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="20889-738">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="20889-738">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="20889-739">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="20889-739">Reply with a body and a file attachment.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="20889-740">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="20889-740">Reply with a body and an item attachment.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="20889-741">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="20889-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="20889-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="20889-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="20889-743">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="20889-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-744">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="20889-744">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-745">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-745">Requirements</span></span>

|<span data-ttu-id="20889-746">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-746">Requirement</span></span>| <span data-ttu-id="20889-747">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-748">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-749">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-749">1.0</span></span>|
|[<span data-ttu-id="20889-750">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-751">ReadItem</span></span>|
|[<span data-ttu-id="20889-752">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-753">Read</span><span class="sxs-lookup"><span data-stu-id="20889-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="20889-754">Retorna:</span><span class="sxs-lookup"><span data-stu-id="20889-754">Returns:</span></span>

<span data-ttu-id="20889-755">Tipo: [Entities](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="20889-755">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="20889-756">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-756">Example</span></span>

<span data-ttu-id="20889-757">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="20889-757">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="20889-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="20889-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="20889-759">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="20889-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-760">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="20889-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="20889-761">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="20889-761">Parameters</span></span>

|<span data-ttu-id="20889-762">Name</span><span class="sxs-lookup"><span data-stu-id="20889-762">Name</span></span>| <span data-ttu-id="20889-763">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-763">Type</span></span>| <span data-ttu-id="20889-764">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="20889-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="20889-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="20889-766">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="20889-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="20889-767">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-767">Requirements</span></span>

|<span data-ttu-id="20889-768">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-768">Requirement</span></span>| <span data-ttu-id="20889-769">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-770">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-771">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-771">1.0</span></span>|
|[<span data-ttu-id="20889-772">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-773">Restrito</span><span class="sxs-lookup"><span data-stu-id="20889-773">Restricted</span></span>|
|[<span data-ttu-id="20889-774">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-775">Read</span><span class="sxs-lookup"><span data-stu-id="20889-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="20889-776">Retorna:</span><span class="sxs-lookup"><span data-stu-id="20889-776">Returns:</span></span>

<span data-ttu-id="20889-777">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="20889-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="20889-778">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="20889-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="20889-779">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="20889-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="20889-780">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="20889-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="20889-781">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="20889-781">Value of `entityType`</span></span> | <span data-ttu-id="20889-782">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="20889-782">Type of objects in returned array</span></span> | <span data-ttu-id="20889-783">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="20889-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="20889-784">String</span><span class="sxs-lookup"><span data-stu-id="20889-784">String</span></span> | <span data-ttu-id="20889-785">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="20889-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="20889-786">Contato</span><span class="sxs-lookup"><span data-stu-id="20889-786">Contact</span></span> | <span data-ttu-id="20889-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="20889-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="20889-788">String</span><span class="sxs-lookup"><span data-stu-id="20889-788">String</span></span> | <span data-ttu-id="20889-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="20889-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="20889-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="20889-790">MeetingSuggestion</span></span> | <span data-ttu-id="20889-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="20889-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="20889-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="20889-792">PhoneNumber</span></span> | <span data-ttu-id="20889-793">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="20889-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="20889-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="20889-794">TaskSuggestion</span></span> | <span data-ttu-id="20889-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="20889-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="20889-796">String</span><span class="sxs-lookup"><span data-stu-id="20889-796">String</span></span> | <span data-ttu-id="20889-797">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="20889-797">**Restricted**</span></span> |

<span data-ttu-id="20889-798">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="20889-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="20889-799">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-799">Example</span></span>

<span data-ttu-id="20889-800">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="20889-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```javascript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="20889-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="20889-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="20889-802">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="20889-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-803">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="20889-803">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="20889-804">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="20889-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="20889-805">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="20889-805">Parameters</span></span>

|<span data-ttu-id="20889-806">Name</span><span class="sxs-lookup"><span data-stu-id="20889-806">Name</span></span>| <span data-ttu-id="20889-807">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-807">Type</span></span>| <span data-ttu-id="20889-808">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="20889-809">String</span><span class="sxs-lookup"><span data-stu-id="20889-809">String</span></span>|<span data-ttu-id="20889-810">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="20889-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="20889-811">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-811">Requirements</span></span>

|<span data-ttu-id="20889-812">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-812">Requirement</span></span>| <span data-ttu-id="20889-813">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-814">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-815">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-815">1.0</span></span>|
|[<span data-ttu-id="20889-816">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-817">ReadItem</span></span>|
|[<span data-ttu-id="20889-818">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="20889-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-819">Read</span><span class="sxs-lookup"><span data-stu-id="20889-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="20889-820">Retorna:</span><span class="sxs-lookup"><span data-stu-id="20889-820">Returns:</span></span>

<span data-ttu-id="20889-p153">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="20889-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="20889-823">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="20889-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="20889-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="20889-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="20889-825">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="20889-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-826">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="20889-826">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="20889-p154">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="20889-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="20889-830">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="20889-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="20889-831">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="20889-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="20889-p155">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="20889-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="20889-835">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-835">Requirements</span></span>

|<span data-ttu-id="20889-836">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-836">Requirement</span></span>| <span data-ttu-id="20889-837">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-838">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-839">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-839">1.0</span></span>|
|[<span data-ttu-id="20889-840">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-841">ReadItem</span></span>|
|[<span data-ttu-id="20889-842">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="20889-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-843">Read</span><span class="sxs-lookup"><span data-stu-id="20889-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="20889-844">Retorna:</span><span class="sxs-lookup"><span data-stu-id="20889-844">Returns:</span></span>

<span data-ttu-id="20889-p156">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="20889-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="20889-847">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="20889-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="20889-848">Object</span><span class="sxs-lookup"><span data-stu-id="20889-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="20889-849">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-849">Example</span></span>

<span data-ttu-id="20889-850">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="20889-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="20889-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="20889-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="20889-852">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="20889-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-853">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="20889-853">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="20889-854">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="20889-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="20889-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="20889-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="20889-857">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="20889-857">Parameters</span></span>

|<span data-ttu-id="20889-858">Name</span><span class="sxs-lookup"><span data-stu-id="20889-858">Name</span></span>| <span data-ttu-id="20889-859">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-859">Type</span></span>| <span data-ttu-id="20889-860">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="20889-861">String</span><span class="sxs-lookup"><span data-stu-id="20889-861">String</span></span>|<span data-ttu-id="20889-862">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="20889-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="20889-863">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-863">Requirements</span></span>

|<span data-ttu-id="20889-864">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-864">Requirement</span></span>| <span data-ttu-id="20889-865">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-866">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-867">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-867">1.0</span></span>|
|[<span data-ttu-id="20889-868">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-869">ReadItem</span></span>|
|[<span data-ttu-id="20889-870">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="20889-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-871">Read</span><span class="sxs-lookup"><span data-stu-id="20889-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="20889-872">Retorna:</span><span class="sxs-lookup"><span data-stu-id="20889-872">Returns:</span></span>

<span data-ttu-id="20889-873">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="20889-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="20889-874">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="20889-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="20889-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="20889-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="20889-876">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-876">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="20889-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="20889-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="20889-878">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="20889-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="20889-p158">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="20889-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="20889-881">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="20889-881">Parameters</span></span>

|<span data-ttu-id="20889-882">Name</span><span class="sxs-lookup"><span data-stu-id="20889-882">Name</span></span>| <span data-ttu-id="20889-883">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-883">Type</span></span>| <span data-ttu-id="20889-884">Atributos</span><span class="sxs-lookup"><span data-stu-id="20889-884">Attributes</span></span>| <span data-ttu-id="20889-885">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="20889-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="20889-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="20889-p159">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="20889-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="20889-890">Object</span><span class="sxs-lookup"><span data-stu-id="20889-890">Object</span></span>| <span data-ttu-id="20889-891">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-891">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-892">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="20889-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="20889-893">Objeto</span><span class="sxs-lookup"><span data-stu-id="20889-893">Object</span></span>| <span data-ttu-id="20889-894">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-894">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-895">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="20889-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="20889-896">function</span><span class="sxs-lookup"><span data-stu-id="20889-896">function</span></span>||<span data-ttu-id="20889-897">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="20889-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="20889-898">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="20889-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="20889-899">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="20889-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="20889-900">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-900">Requirements</span></span>

|<span data-ttu-id="20889-901">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-901">Requirement</span></span>| <span data-ttu-id="20889-902">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-903">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-904">1.2</span><span class="sxs-lookup"><span data-stu-id="20889-904">1.2</span></span>|
|[<span data-ttu-id="20889-905">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="20889-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="20889-907">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-908">Escrever</span><span class="sxs-lookup"><span data-stu-id="20889-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="20889-909">Retorna:</span><span class="sxs-lookup"><span data-stu-id="20889-909">Returns:</span></span>

<span data-ttu-id="20889-910">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="20889-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="20889-911">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="20889-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="20889-912">String</span><span class="sxs-lookup"><span data-stu-id="20889-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="20889-913">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-913">Example</span></span>

```javascript
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
  // Check for errors.
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="20889-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="20889-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="20889-915">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="20889-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="20889-p161">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="20889-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="20889-919">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="20889-919">Parameters</span></span>

|<span data-ttu-id="20889-920">Name</span><span class="sxs-lookup"><span data-stu-id="20889-920">Name</span></span>| <span data-ttu-id="20889-921">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-921">Type</span></span>| <span data-ttu-id="20889-922">Atributos</span><span class="sxs-lookup"><span data-stu-id="20889-922">Attributes</span></span>| <span data-ttu-id="20889-923">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="20889-924">function</span><span class="sxs-lookup"><span data-stu-id="20889-924">function</span></span>||<span data-ttu-id="20889-925">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="20889-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="20889-926">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="20889-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="20889-927">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="20889-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="20889-928">Objeto</span><span class="sxs-lookup"><span data-stu-id="20889-928">Object</span></span>| <span data-ttu-id="20889-929">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-929">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-930">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="20889-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="20889-931">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="20889-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="20889-932">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-932">Requirements</span></span>

|<span data-ttu-id="20889-933">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-933">Requirement</span></span>| <span data-ttu-id="20889-934">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-935">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-936">1.0</span><span class="sxs-lookup"><span data-stu-id="20889-936">1.0</span></span>|
|[<span data-ttu-id="20889-937">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20889-938">ReadItem</span></span>|
|[<span data-ttu-id="20889-939">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="20889-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-940">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="20889-940">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-941">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-941">Example</span></span>

<span data-ttu-id="20889-p164">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="20889-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="20889-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="20889-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="20889-946">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="20889-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="20889-p165">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="20889-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="20889-951">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="20889-951">Parameters</span></span>

|<span data-ttu-id="20889-952">Name</span><span class="sxs-lookup"><span data-stu-id="20889-952">Name</span></span>| <span data-ttu-id="20889-953">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-953">Type</span></span>| <span data-ttu-id="20889-954">Atributos</span><span class="sxs-lookup"><span data-stu-id="20889-954">Attributes</span></span>| <span data-ttu-id="20889-955">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="20889-956">String</span><span class="sxs-lookup"><span data-stu-id="20889-956">String</span></span>||<span data-ttu-id="20889-957">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="20889-957">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="20889-958">Object</span><span class="sxs-lookup"><span data-stu-id="20889-958">Object</span></span>| <span data-ttu-id="20889-959">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-959">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-960">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="20889-960">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="20889-961">Object</span><span class="sxs-lookup"><span data-stu-id="20889-961">Object</span></span>| <span data-ttu-id="20889-962">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-962">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-963">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="20889-963">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="20889-964">function</span><span class="sxs-lookup"><span data-stu-id="20889-964">function</span></span>| <span data-ttu-id="20889-965">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-965">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-966">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="20889-966">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="20889-967">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="20889-967">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="20889-968">Erros</span><span class="sxs-lookup"><span data-stu-id="20889-968">Errors</span></span>

| <span data-ttu-id="20889-969">Código de erro</span><span class="sxs-lookup"><span data-stu-id="20889-969">Error code</span></span> | <span data-ttu-id="20889-970">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-970">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="20889-971">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="20889-971">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="20889-972">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-972">Requirements</span></span>

|<span data-ttu-id="20889-973">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-973">Requirement</span></span>| <span data-ttu-id="20889-974">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-974">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-975">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-975">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-976">1.1</span><span class="sxs-lookup"><span data-stu-id="20889-976">1.1</span></span>|
|[<span data-ttu-id="20889-977">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-977">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-978">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="20889-978">ReadWriteItem</span></span>|
|[<span data-ttu-id="20889-979">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-979">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-980">Escrever</span><span class="sxs-lookup"><span data-stu-id="20889-980">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-981">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-981">Example</span></span>

<span data-ttu-id="20889-982">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="20889-982">The following code removes an attachment with an identifier of '0'.</span></span>

```javascript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="20889-983">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="20889-983">saveAsync([options], callback)</span></span>

<span data-ttu-id="20889-984">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="20889-984">Asynchronously saves an item.</span></span>

<span data-ttu-id="20889-p166">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="20889-p166">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-988">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="20889-988">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="20889-989">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="20889-989">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="20889-p168">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="20889-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="20889-993">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="20889-993">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="20889-994">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="20889-994">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="20889-995">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="20889-995">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="20889-996">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="20889-996">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="20889-997">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="20889-997">Parameters</span></span>

|<span data-ttu-id="20889-998">Name</span><span class="sxs-lookup"><span data-stu-id="20889-998">Name</span></span>| <span data-ttu-id="20889-999">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-999">Type</span></span>| <span data-ttu-id="20889-1000">Atributos</span><span class="sxs-lookup"><span data-stu-id="20889-1000">Attributes</span></span>| <span data-ttu-id="20889-1001">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-1001">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="20889-1002">Object</span><span class="sxs-lookup"><span data-stu-id="20889-1002">Object</span></span>| <span data-ttu-id="20889-1003">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-1003">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-1004">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="20889-1004">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="20889-1005">Object</span><span class="sxs-lookup"><span data-stu-id="20889-1005">Object</span></span>| <span data-ttu-id="20889-1006">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-1007">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="20889-1007">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="20889-1008">function</span><span class="sxs-lookup"><span data-stu-id="20889-1008">function</span></span>||<span data-ttu-id="20889-1009">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="20889-1009">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="20889-1010">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="20889-1010">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="20889-1011">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-1011">Requirements</span></span>

|<span data-ttu-id="20889-1012">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-1012">Requirement</span></span>| <span data-ttu-id="20889-1013">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-1013">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-1014">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-1014">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-1015">1.3</span><span class="sxs-lookup"><span data-stu-id="20889-1015">1.3</span></span>|
|[<span data-ttu-id="20889-1016">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-1016">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-1017">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="20889-1017">ReadWriteItem</span></span>|
|[<span data-ttu-id="20889-1018">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-1018">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-1019">Escrever</span><span class="sxs-lookup"><span data-stu-id="20889-1019">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="20889-1020">Exemplos</span><span class="sxs-lookup"><span data-stu-id="20889-1020">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="20889-p170">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="20889-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="20889-1023">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="20889-1023">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="20889-1024">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="20889-1024">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="20889-p171">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="20889-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="20889-1028">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="20889-1028">Parameters</span></span>

|<span data-ttu-id="20889-1029">Name</span><span class="sxs-lookup"><span data-stu-id="20889-1029">Name</span></span>| <span data-ttu-id="20889-1030">Tipo</span><span class="sxs-lookup"><span data-stu-id="20889-1030">Type</span></span>| <span data-ttu-id="20889-1031">Atributos</span><span class="sxs-lookup"><span data-stu-id="20889-1031">Attributes</span></span>| <span data-ttu-id="20889-1032">Descrição</span><span class="sxs-lookup"><span data-stu-id="20889-1032">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="20889-1033">String</span><span class="sxs-lookup"><span data-stu-id="20889-1033">String</span></span>||<span data-ttu-id="20889-p172">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="20889-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="20889-1037">Object</span><span class="sxs-lookup"><span data-stu-id="20889-1037">Object</span></span>| <span data-ttu-id="20889-1038">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-1038">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-1039">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="20889-1039">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="20889-1040">Object</span><span class="sxs-lookup"><span data-stu-id="20889-1040">Object</span></span>| <span data-ttu-id="20889-1041">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-1041">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-1042">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="20889-1042">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="20889-1043">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="20889-1043">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="20889-1044">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="20889-1044">&lt;optional&gt;</span></span>|<span data-ttu-id="20889-p173">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="20889-p173">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="20889-p174">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="20889-p174">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="20889-1049">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="20889-1049">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="20889-1050">function</span><span class="sxs-lookup"><span data-stu-id="20889-1050">function</span></span>||<span data-ttu-id="20889-1051">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="20889-1051">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="20889-1052">Requisitos</span><span class="sxs-lookup"><span data-stu-id="20889-1052">Requirements</span></span>

|<span data-ttu-id="20889-1053">Requisito</span><span class="sxs-lookup"><span data-stu-id="20889-1053">Requirement</span></span>| <span data-ttu-id="20889-1054">Valor</span><span class="sxs-lookup"><span data-stu-id="20889-1054">Value</span></span>|
|---|---|
|[<span data-ttu-id="20889-1055">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="20889-1055">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20889-1056">1.2</span><span class="sxs-lookup"><span data-stu-id="20889-1056">1.2</span></span>|
|[<span data-ttu-id="20889-1057">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="20889-1057">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20889-1058">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="20889-1058">ReadWriteItem</span></span>|
|[<span data-ttu-id="20889-1059">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="20889-1059">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20889-1060">Escrever</span><span class="sxs-lookup"><span data-stu-id="20889-1060">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="20889-1061">Exemplo</span><span class="sxs-lookup"><span data-stu-id="20889-1061">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
