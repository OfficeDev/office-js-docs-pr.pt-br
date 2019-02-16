---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: c0b956cac0410ef7d8e8e0d59a69e221e29c540a
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068039"
---
# <a name="item"></a><span data-ttu-id="1d373-102">item</span><span class="sxs-lookup"><span data-stu-id="1d373-102">item</span></span>

### <span data-ttu-id="1d373-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="1d373-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="1d373-p102">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="1d373-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-107">Requirements</span></span>

|<span data-ttu-id="1d373-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-108">Requirement</span></span>| <span data-ttu-id="1d373-109">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-111">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-111">1.0</span></span>|
|[<span data-ttu-id="1d373-112">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-113">Restrito</span><span class="sxs-lookup"><span data-stu-id="1d373-113">Restricted</span></span>|
|[<span data-ttu-id="1d373-114">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-115">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="1d373-116">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-116">Example</span></span>

<span data-ttu-id="1d373-117">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="1d373-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="1d373-118">Membros</span><span class="sxs-lookup"><span data-stu-id="1d373-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="1d373-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1d373-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="1d373-p103">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d373-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1d373-122">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="1d373-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="1d373-123">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="1d373-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-124">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-124">Type</span></span>

*   <span data-ttu-id="1d373-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1d373-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-126">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-126">Requirements</span></span>

|<span data-ttu-id="1d373-127">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-127">Requirement</span></span>| <span data-ttu-id="1d373-128">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-129">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-130">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-130">1.0</span></span>|
|[<span data-ttu-id="1d373-131">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-132">ReadItem</span></span>|
|[<span data-ttu-id="1d373-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-134">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-135">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-135">Example</span></span>

<span data-ttu-id="1d373-136">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="1d373-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="1d373-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d373-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="1d373-138">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1d373-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="1d373-139">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="1d373-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-140">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-140">Type</span></span>

*   [<span data-ttu-id="1d373-141">Destinatários</span><span class="sxs-lookup"><span data-stu-id="1d373-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="1d373-142">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-142">Requirements</span></span>

|<span data-ttu-id="1d373-143">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-143">Requirement</span></span>| <span data-ttu-id="1d373-144">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-145">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-146">1.1</span><span class="sxs-lookup"><span data-stu-id="1d373-146">1.1</span></span>|
|[<span data-ttu-id="1d373-147">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-148">ReadItem</span></span>|
|[<span data-ttu-id="1d373-149">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-150">Escrever</span><span class="sxs-lookup"><span data-stu-id="1d373-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-151">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="1d373-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="1d373-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="1d373-153">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="1d373-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-154">Type</span></span>

*   [<span data-ttu-id="1d373-155">Corpo</span><span class="sxs-lookup"><span data-stu-id="1d373-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="1d373-156">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-156">Requirements</span></span>

|<span data-ttu-id="1d373-157">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-157">Requirement</span></span>| <span data-ttu-id="1d373-158">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-159">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-160">1.1</span><span class="sxs-lookup"><span data-stu-id="1d373-160">1.1</span></span>|
|[<span data-ttu-id="1d373-161">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-162">ReadItem</span></span>|
|[<span data-ttu-id="1d373-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-164">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-165">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-165">Example</span></span>

<span data-ttu-id="1d373-166">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="1d373-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="1d373-167">A seguir apresentamos um exemplo do parâmetro result passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1d373-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="1d373-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d373-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="1d373-169">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1d373-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="1d373-170">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1d373-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d373-171">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-171">Read mode</span></span>

<span data-ttu-id="1d373-p107">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="1d373-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="1d373-174">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1d373-174">Compose mode</span></span>

<span data-ttu-id="1d373-175">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="1d373-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1d373-176">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-176">Type</span></span>

*   <span data-ttu-id="1d373-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d373-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-178">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-178">Requirements</span></span>

|<span data-ttu-id="1d373-179">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-179">Requirement</span></span>| <span data-ttu-id="1d373-180">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-181">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-182">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-182">1.0</span></span>|
|[<span data-ttu-id="1d373-183">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-183">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-184">ReadItem</span></span>|
|[<span data-ttu-id="1d373-185">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-185">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-186">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-186">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="1d373-187">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="1d373-187">(nullable) conversationId :String</span></span>

<span data-ttu-id="1d373-188">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="1d373-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="1d373-p108">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="1d373-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="1d373-p109">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="1d373-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-193">Type</span></span>

*   <span data-ttu-id="1d373-194">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d373-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-195">Requirements</span></span>

|<span data-ttu-id="1d373-196">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-196">Requirement</span></span>| <span data-ttu-id="1d373-197">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-199">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-199">1.0</span></span>|
|[<span data-ttu-id="1d373-200">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-200">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-201">ReadItem</span></span>|
|[<span data-ttu-id="1d373-202">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-202">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-203">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-204">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="1d373-205">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="1d373-205">dateTimeCreated :Date</span></span>

<span data-ttu-id="1d373-p110">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d373-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-208">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-208">Type</span></span>

*   <span data-ttu-id="1d373-209">Data</span><span class="sxs-lookup"><span data-stu-id="1d373-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-210">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-210">Requirements</span></span>

|<span data-ttu-id="1d373-211">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-211">Requirement</span></span>| <span data-ttu-id="1d373-212">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-213">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-214">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-214">1.0</span></span>|
|[<span data-ttu-id="1d373-215">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-215">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-216">ReadItem</span></span>|
|[<span data-ttu-id="1d373-217">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-217">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-218">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-219">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="1d373-220">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="1d373-220">dateTimeModified :Date</span></span>

<span data-ttu-id="1d373-p111">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d373-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1d373-223">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d373-223">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-224">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-224">Type</span></span>

*   <span data-ttu-id="1d373-225">Data</span><span class="sxs-lookup"><span data-stu-id="1d373-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-226">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-226">Requirements</span></span>

|<span data-ttu-id="1d373-227">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-227">Requirement</span></span>| <span data-ttu-id="1d373-228">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-229">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-230">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-230">1.0</span></span>|
|[<span data-ttu-id="1d373-231">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-231">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-232">ReadItem</span></span>|
|[<span data-ttu-id="1d373-233">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-233">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-234">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-235">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="1d373-236">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="1d373-236">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="1d373-237">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="1d373-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="1d373-p112">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="1d373-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d373-240">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-240">Read mode</span></span>

<span data-ttu-id="1d373-241">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="1d373-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="1d373-242">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1d373-242">Compose mode</span></span>

<span data-ttu-id="1d373-243">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="1d373-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="1d373-244">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="1d373-244">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="1d373-245">O exemplo a seguir define a hora de término de um compromisso usando [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) o método do `Time` objeto.</span><span class="sxs-lookup"><span data-stu-id="1d373-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1d373-246">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-246">Type</span></span>

*   <span data-ttu-id="1d373-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="1d373-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-248">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-248">Requirements</span></span>

|<span data-ttu-id="1d373-249">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-249">Requirement</span></span>| <span data-ttu-id="1d373-250">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-251">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-252">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-252">1.0</span></span>|
|[<span data-ttu-id="1d373-253">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-253">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-254">ReadItem</span></span>|
|[<span data-ttu-id="1d373-255">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-255">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-256">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="1d373-257">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1d373-257">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="1d373-p113">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d373-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="1d373-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="1d373-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1d373-262">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="1d373-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-263">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-263">Type</span></span>

*   [<span data-ttu-id="1d373-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1d373-264">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1d373-265">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-265">Requirements</span></span>

|<span data-ttu-id="1d373-266">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-266">Requirement</span></span>| <span data-ttu-id="1d373-267">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-268">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-269">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-269">1.0</span></span>|
|[<span data-ttu-id="1d373-270">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-270">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-271">ReadItem</span></span>|
|[<span data-ttu-id="1d373-272">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-272">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-273">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-274">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="1d373-275">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="1d373-275">internetMessageId :String</span></span>

<span data-ttu-id="1d373-p115">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d373-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-278">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-278">Type</span></span>

*   <span data-ttu-id="1d373-279">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d373-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-280">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-280">Requirements</span></span>

|<span data-ttu-id="1d373-281">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-281">Requirement</span></span>| <span data-ttu-id="1d373-282">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-283">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-284">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-284">1.0</span></span>|
|[<span data-ttu-id="1d373-285">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-285">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-286">ReadItem</span></span>|
|[<span data-ttu-id="1d373-287">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-287">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-288">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-289">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="1d373-290">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="1d373-290">itemClass :String</span></span>

<span data-ttu-id="1d373-p116">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d373-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="1d373-p117">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="1d373-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="1d373-295">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-295">Type</span></span> | <span data-ttu-id="1d373-296">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-296">Description</span></span> | <span data-ttu-id="1d373-297">classe de item</span><span class="sxs-lookup"><span data-stu-id="1d373-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="1d373-298">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="1d373-298">Appointment items</span></span> | <span data-ttu-id="1d373-299">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="1d373-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="1d373-300">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="1d373-300">Message items</span></span> | <span data-ttu-id="1d373-301">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="1d373-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="1d373-302">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="1d373-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-303">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-303">Type</span></span>

*   <span data-ttu-id="1d373-304">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d373-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-305">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-305">Requirements</span></span>

|<span data-ttu-id="1d373-306">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-306">Requirement</span></span>| <span data-ttu-id="1d373-307">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-308">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-309">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-309">1.0</span></span>|
|[<span data-ttu-id="1d373-310">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-310">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-311">ReadItem</span></span>|
|[<span data-ttu-id="1d373-312">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-312">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-313">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-314">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="1d373-315">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="1d373-315">(nullable) itemId :String</span></span>

<span data-ttu-id="1d373-p118">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d373-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1d373-318">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="1d373-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1d373-319">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="1d373-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1d373-320">Antes de fazer chamadas API REST usando esse valor, ele deve ser convertido usando `Office.context.mailbox.convertToRestId`, que está disponível a partir do conjunto de requisitos 1.3.</span><span class="sxs-lookup"><span data-stu-id="1d373-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="1d373-321">Para saber mais, consulte [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="1d373-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-322">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-322">Type</span></span>

*   <span data-ttu-id="1d373-323">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d373-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-324">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-324">Requirements</span></span>

|<span data-ttu-id="1d373-325">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-325">Requirement</span></span>| <span data-ttu-id="1d373-326">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-327">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-328">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-328">1.0</span></span>|
|[<span data-ttu-id="1d373-329">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-330">ReadItem</span></span>|
|[<span data-ttu-id="1d373-331">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-332">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-333">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-333">Example</span></span>

<span data-ttu-id="1d373-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="1d373-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="1d373-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="1d373-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="1d373-337">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="1d373-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="1d373-338">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="1d373-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-339">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-339">Type</span></span>

*   [<span data-ttu-id="1d373-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="1d373-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="1d373-341">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-341">Requirements</span></span>

|<span data-ttu-id="1d373-342">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-342">Requirement</span></span>| <span data-ttu-id="1d373-343">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-344">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-345">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-345">1.0</span></span>|
|[<span data-ttu-id="1d373-346">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-346">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-347">ReadItem</span></span>|
|[<span data-ttu-id="1d373-348">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-348">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-349">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-350">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="1d373-351">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="1d373-351">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="1d373-352">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="1d373-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d373-353">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-353">Read mode</span></span>

<span data-ttu-id="1d373-354">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="1d373-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="1d373-355">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1d373-355">Compose mode</span></span>

<span data-ttu-id="1d373-356">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="1d373-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1d373-357">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-357">Type</span></span>

*   <span data-ttu-id="1d373-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="1d373-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-359">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-359">Requirements</span></span>

|<span data-ttu-id="1d373-360">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-360">Requirement</span></span>| <span data-ttu-id="1d373-361">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-362">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-363">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-363">1.0</span></span>|
|[<span data-ttu-id="1d373-364">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-365">ReadItem</span></span>|
|[<span data-ttu-id="1d373-366">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-367">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="1d373-368">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="1d373-368">normalizedSubject :String</span></span>

<span data-ttu-id="1d373-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d373-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="1d373-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject).</span><span class="sxs-lookup"><span data-stu-id="1d373-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-373">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-373">Type</span></span>

*   <span data-ttu-id="1d373-374">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d373-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-375">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-375">Requirements</span></span>

|<span data-ttu-id="1d373-376">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-376">Requirement</span></span>| <span data-ttu-id="1d373-377">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-378">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-379">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-379">1.0</span></span>|
|[<span data-ttu-id="1d373-380">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-380">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-381">ReadItem</span></span>|
|[<span data-ttu-id="1d373-382">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-382">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-383">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-384">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="1d373-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d373-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="1d373-386">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="1d373-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="1d373-387">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1d373-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d373-388">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-388">Read mode</span></span>

<span data-ttu-id="1d373-389">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="1d373-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1d373-390">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1d373-390">Compose mode</span></span>

<span data-ttu-id="1d373-391">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="1d373-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1d373-392">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-392">Type</span></span>

*   <span data-ttu-id="1d373-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d373-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-394">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-394">Requirements</span></span>

|<span data-ttu-id="1d373-395">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-395">Requirement</span></span>| <span data-ttu-id="1d373-396">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-397">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-398">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-398">1.0</span></span>|
|[<span data-ttu-id="1d373-399">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-399">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-400">ReadItem</span></span>|
|[<span data-ttu-id="1d373-401">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-401">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-402">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="1d373-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1d373-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="1d373-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d373-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-406">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-406">Type</span></span>

*   [<span data-ttu-id="1d373-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1d373-407">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1d373-408">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-408">Requirements</span></span>

|<span data-ttu-id="1d373-409">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-409">Requirement</span></span>| <span data-ttu-id="1d373-410">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-411">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-412">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-412">1.0</span></span>|
|[<span data-ttu-id="1d373-413">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-414">ReadItem</span></span>|
|[<span data-ttu-id="1d373-415">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-416">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-417">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="1d373-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d373-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="1d373-419">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="1d373-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="1d373-420">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1d373-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d373-421">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-421">Read mode</span></span>

<span data-ttu-id="1d373-422">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="1d373-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1d373-423">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1d373-423">Compose mode</span></span>

<span data-ttu-id="1d373-424">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="1d373-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="1d373-425">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-425">Type</span></span>

*   <span data-ttu-id="1d373-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d373-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-427">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-427">Requirements</span></span>

|<span data-ttu-id="1d373-428">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-428">Requirement</span></span>| <span data-ttu-id="1d373-429">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-430">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-431">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-431">1.0</span></span>|
|[<span data-ttu-id="1d373-432">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-432">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-433">ReadItem</span></span>|
|[<span data-ttu-id="1d373-434">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-434">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-435">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="1d373-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1d373-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="1d373-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d373-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="1d373-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="1d373-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1d373-441">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="1d373-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1d373-442">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-442">Type</span></span>

*   [<span data-ttu-id="1d373-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1d373-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1d373-444">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-444">Requirements</span></span>

|<span data-ttu-id="1d373-445">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-445">Requirement</span></span>| <span data-ttu-id="1d373-446">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-447">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-448">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-448">1.0</span></span>|
|[<span data-ttu-id="1d373-449">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-450">ReadItem</span></span>|
|[<span data-ttu-id="1d373-451">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-452">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-453">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="1d373-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="1d373-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="1d373-455">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="1d373-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="1d373-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="1d373-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d373-458">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-458">Read mode</span></span>

<span data-ttu-id="1d373-459">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="1d373-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="1d373-460">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1d373-460">Compose mode</span></span>

<span data-ttu-id="1d373-461">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="1d373-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="1d373-462">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="1d373-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="1d373-463">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="1d373-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1d373-464">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-464">Type</span></span>

*   <span data-ttu-id="1d373-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="1d373-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-466">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-466">Requirements</span></span>

|<span data-ttu-id="1d373-467">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-467">Requirement</span></span>| <span data-ttu-id="1d373-468">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-469">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-470">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-470">1.0</span></span>|
|[<span data-ttu-id="1d373-471">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-472">ReadItem</span></span>|
|[<span data-ttu-id="1d373-473">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-474">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-474">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="1d373-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1d373-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="1d373-476">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="1d373-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="1d373-477">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="1d373-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d373-478">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-478">Read mode</span></span>

<span data-ttu-id="1d373-p130">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="1d373-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="1d373-481">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1d373-481">Compose mode</span></span>

<span data-ttu-id="1d373-482">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="1d373-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="1d373-483">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-483">Type</span></span>

*   <span data-ttu-id="1d373-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1d373-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-485">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-485">Requirements</span></span>

|<span data-ttu-id="1d373-486">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-486">Requirement</span></span>| <span data-ttu-id="1d373-487">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-488">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-489">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-489">1.0</span></span>|
|[<span data-ttu-id="1d373-490">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-491">ReadItem</span></span>|
|[<span data-ttu-id="1d373-492">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-493">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-493">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="1d373-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d373-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="1d373-495">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1d373-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="1d373-496">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1d373-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d373-497">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-497">Read mode</span></span>

<span data-ttu-id="1d373-p132">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="1d373-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="1d373-500">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1d373-500">Compose mode</span></span>

<span data-ttu-id="1d373-501">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="1d373-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1d373-502">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-502">Type</span></span>

*   <span data-ttu-id="1d373-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d373-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-504">Requirements</span></span>

|<span data-ttu-id="1d373-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-505">Requirement</span></span>| <span data-ttu-id="1d373-506">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-507">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-508">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-508">1.0</span></span>|
|[<span data-ttu-id="1d373-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-510">ReadItem</span></span>|
|[<span data-ttu-id="1d373-511">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-512">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="1d373-513">Métodos</span><span class="sxs-lookup"><span data-stu-id="1d373-513">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="1d373-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1d373-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1d373-515">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="1d373-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1d373-516">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="1d373-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="1d373-517">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="1d373-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d373-518">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="1d373-518">Parameters</span></span>

|<span data-ttu-id="1d373-519">Nome</span><span class="sxs-lookup"><span data-stu-id="1d373-519">Name</span></span>| <span data-ttu-id="1d373-520">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-520">Type</span></span>| <span data-ttu-id="1d373-521">Atributos</span><span class="sxs-lookup"><span data-stu-id="1d373-521">Attributes</span></span>| <span data-ttu-id="1d373-522">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="1d373-523">String</span><span class="sxs-lookup"><span data-stu-id="1d373-523">String</span></span>||<span data-ttu-id="1d373-p133">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1d373-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1d373-526">String</span><span class="sxs-lookup"><span data-stu-id="1d373-526">String</span></span>||<span data-ttu-id="1d373-p134">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1d373-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1d373-529">Object</span><span class="sxs-lookup"><span data-stu-id="1d373-529">Object</span></span>| <span data-ttu-id="1d373-530">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-530">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-531">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1d373-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1d373-532">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d373-532">Object</span></span>| <span data-ttu-id="1d373-533">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-533">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-534">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1d373-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1d373-535">function</span><span class="sxs-lookup"><span data-stu-id="1d373-535">function</span></span>| <span data-ttu-id="1d373-536">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-536">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-537">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d373-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1d373-538">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d373-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1d373-539">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="1d373-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1d373-540">Erros</span><span class="sxs-lookup"><span data-stu-id="1d373-540">Errors</span></span>

| <span data-ttu-id="1d373-541">Código de erro</span><span class="sxs-lookup"><span data-stu-id="1d373-541">Error code</span></span> | <span data-ttu-id="1d373-542">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="1d373-543">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="1d373-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="1d373-544">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="1d373-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1d373-545">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="1d373-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1d373-546">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-546">Requirements</span></span>

|<span data-ttu-id="1d373-547">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-547">Requirement</span></span>| <span data-ttu-id="1d373-548">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-549">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-550">1.1</span><span class="sxs-lookup"><span data-stu-id="1d373-550">1.1</span></span>|
|[<span data-ttu-id="1d373-551">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-551">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1d373-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="1d373-553">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-554">Escrever</span><span class="sxs-lookup"><span data-stu-id="1d373-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-555">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-555">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="1d373-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1d373-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1d373-557">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="1d373-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="1d373-p135">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="1d373-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="1d373-561">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="1d373-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="1d373-562">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="1d373-562">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d373-563">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="1d373-563">Parameters</span></span>

|<span data-ttu-id="1d373-564">Nome</span><span class="sxs-lookup"><span data-stu-id="1d373-564">Name</span></span>| <span data-ttu-id="1d373-565">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-565">Type</span></span>| <span data-ttu-id="1d373-566">Atributos</span><span class="sxs-lookup"><span data-stu-id="1d373-566">Attributes</span></span>| <span data-ttu-id="1d373-567">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="1d373-568">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d373-568">String</span></span>||<span data-ttu-id="1d373-p136">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1d373-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1d373-571">String</span><span class="sxs-lookup"><span data-stu-id="1d373-571">String</span></span>||<span data-ttu-id="1d373-572">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="1d373-572">The subject of the item to be attached.</span></span> <span data-ttu-id="1d373-573">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1d373-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1d373-574">Object</span><span class="sxs-lookup"><span data-stu-id="1d373-574">Object</span></span>| <span data-ttu-id="1d373-575">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-575">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-576">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1d373-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1d373-577">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d373-577">Object</span></span>| <span data-ttu-id="1d373-578">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-578">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-579">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1d373-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1d373-580">function</span><span class="sxs-lookup"><span data-stu-id="1d373-580">function</span></span>| <span data-ttu-id="1d373-581">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-581">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-582">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d373-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1d373-583">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d373-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1d373-584">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="1d373-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1d373-585">Erros</span><span class="sxs-lookup"><span data-stu-id="1d373-585">Errors</span></span>

| <span data-ttu-id="1d373-586">Código de erro</span><span class="sxs-lookup"><span data-stu-id="1d373-586">Error code</span></span> | <span data-ttu-id="1d373-587">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1d373-588">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="1d373-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1d373-589">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-589">Requirements</span></span>

|<span data-ttu-id="1d373-590">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-590">Requirement</span></span>| <span data-ttu-id="1d373-591">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-592">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-593">1.1</span><span class="sxs-lookup"><span data-stu-id="1d373-593">1.1</span></span>|
|[<span data-ttu-id="1d373-594">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-594">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1d373-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="1d373-596">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-596">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-597">Escrever</span><span class="sxs-lookup"><span data-stu-id="1d373-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-598">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-598">Example</span></span>

<span data-ttu-id="1d373-599">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="1d373-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="1d373-600">displayReplyAllForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1d373-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="1d373-601">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="1d373-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1d373-602">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d373-602">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d373-603">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="1d373-603">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1d373-604">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="1d373-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="1d373-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="1d373-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d373-608">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="1d373-608">Parameters</span></span>

|<span data-ttu-id="1d373-609">Nome</span><span class="sxs-lookup"><span data-stu-id="1d373-609">Name</span></span>| <span data-ttu-id="1d373-610">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-610">Type</span></span>| <span data-ttu-id="1d373-611">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="1d373-612">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1d373-612">String &#124; Object</span></span>| |<span data-ttu-id="1d373-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1d373-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1d373-615">**OU**</span><span class="sxs-lookup"><span data-stu-id="1d373-615">**OR**</span></span><br/><span data-ttu-id="1d373-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="1d373-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1d373-618">String</span><span class="sxs-lookup"><span data-stu-id="1d373-618">String</span></span> | <span data-ttu-id="1d373-619">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-619">&lt;optional&gt;</span></span> | <span data-ttu-id="1d373-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1d373-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1d373-622">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-622">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1d373-623">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-623">&lt;optional&gt;</span></span> | <span data-ttu-id="1d373-624">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="1d373-624">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1d373-625">String</span><span class="sxs-lookup"><span data-stu-id="1d373-625">String</span></span> | | <span data-ttu-id="1d373-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="1d373-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1d373-628">String</span><span class="sxs-lookup"><span data-stu-id="1d373-628">String</span></span> | | <span data-ttu-id="1d373-629">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="1d373-629">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1d373-630">String</span><span class="sxs-lookup"><span data-stu-id="1d373-630">String</span></span> | | <span data-ttu-id="1d373-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="1d373-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1d373-633">String</span><span class="sxs-lookup"><span data-stu-id="1d373-633">String</span></span> | | <span data-ttu-id="1d373-p144">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1d373-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1d373-637">function</span><span class="sxs-lookup"><span data-stu-id="1d373-637">function</span></span> | <span data-ttu-id="1d373-638">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-638">&lt;optional&gt;</span></span> | <span data-ttu-id="1d373-639">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d373-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1d373-640">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-640">Requirements</span></span>

|<span data-ttu-id="1d373-641">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-641">Requirement</span></span>| <span data-ttu-id="1d373-642">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-643">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-644">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-644">1.0</span></span>|
|[<span data-ttu-id="1d373-645">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-645">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-646">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-646">ReadItem</span></span>|
|[<span data-ttu-id="1d373-647">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-647">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-648">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-648">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1d373-649">Exemplos</span><span class="sxs-lookup"><span data-stu-id="1d373-649">Examples</span></span>

<span data-ttu-id="1d373-650">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="1d373-650">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="1d373-651">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="1d373-651">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="1d373-652">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="1d373-652">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1d373-653">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="1d373-653">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1d373-654">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="1d373-654">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1d373-655">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1d373-655">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="1d373-656">displayReplyForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1d373-656">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="1d373-657">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="1d373-657">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1d373-658">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d373-658">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d373-659">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="1d373-659">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1d373-660">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="1d373-660">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="1d373-p145">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="1d373-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d373-664">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="1d373-664">Parameters</span></span>

|<span data-ttu-id="1d373-665">Nome</span><span class="sxs-lookup"><span data-stu-id="1d373-665">Name</span></span>| <span data-ttu-id="1d373-666">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-666">Type</span></span>| <span data-ttu-id="1d373-667">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-667">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="1d373-668">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1d373-668">String &#124; Object</span></span>| | <span data-ttu-id="1d373-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1d373-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1d373-671">**OU**</span><span class="sxs-lookup"><span data-stu-id="1d373-671">**OR**</span></span><br/><span data-ttu-id="1d373-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="1d373-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1d373-674">String</span><span class="sxs-lookup"><span data-stu-id="1d373-674">String</span></span> | <span data-ttu-id="1d373-675">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-675">&lt;optional&gt;</span></span> | <span data-ttu-id="1d373-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1d373-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1d373-678">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-678">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1d373-679">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-679">&lt;optional&gt;</span></span> | <span data-ttu-id="1d373-680">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="1d373-680">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1d373-681">String</span><span class="sxs-lookup"><span data-stu-id="1d373-681">String</span></span> | | <span data-ttu-id="1d373-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="1d373-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1d373-684">String</span><span class="sxs-lookup"><span data-stu-id="1d373-684">String</span></span> | | <span data-ttu-id="1d373-685">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="1d373-685">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1d373-686">String</span><span class="sxs-lookup"><span data-stu-id="1d373-686">String</span></span> | | <span data-ttu-id="1d373-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="1d373-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1d373-689">String</span><span class="sxs-lookup"><span data-stu-id="1d373-689">String</span></span> | | <span data-ttu-id="1d373-p151">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1d373-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1d373-693">function</span><span class="sxs-lookup"><span data-stu-id="1d373-693">function</span></span> | <span data-ttu-id="1d373-694">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-694">&lt;optional&gt;</span></span> | <span data-ttu-id="1d373-695">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d373-695">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1d373-696">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-696">Requirements</span></span>

|<span data-ttu-id="1d373-697">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-697">Requirement</span></span>| <span data-ttu-id="1d373-698">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-699">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-700">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-700">1.0</span></span>|
|[<span data-ttu-id="1d373-701">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-701">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-702">ReadItem</span></span>|
|[<span data-ttu-id="1d373-703">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-703">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-704">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-704">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1d373-705">Exemplos</span><span class="sxs-lookup"><span data-stu-id="1d373-705">Examples</span></span>

<span data-ttu-id="1d373-706">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="1d373-706">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="1d373-707">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="1d373-707">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="1d373-708">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="1d373-708">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1d373-709">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="1d373-709">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1d373-710">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="1d373-710">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1d373-711">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1d373-711">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="1d373-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1d373-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="1d373-713">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="1d373-713">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1d373-714">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d373-714">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-715">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-715">Requirements</span></span>

|<span data-ttu-id="1d373-716">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-716">Requirement</span></span>| <span data-ttu-id="1d373-717">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-717">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-718">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-718">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-719">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-719">1.0</span></span>|
|[<span data-ttu-id="1d373-720">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-720">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-721">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-721">ReadItem</span></span>|
|[<span data-ttu-id="1d373-722">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-722">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-723">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-723">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d373-724">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1d373-724">Returns:</span></span>

<span data-ttu-id="1d373-725">Tipo: [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1d373-725">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1d373-726">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-726">Example</span></span>

<span data-ttu-id="1d373-727">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1d373-727">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="1d373-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1d373-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1d373-729">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="1d373-729">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1d373-730">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d373-730">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d373-731">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="1d373-731">Parameters</span></span>

|<span data-ttu-id="1d373-732">Nome</span><span class="sxs-lookup"><span data-stu-id="1d373-732">Name</span></span>| <span data-ttu-id="1d373-733">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-733">Type</span></span>| <span data-ttu-id="1d373-734">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-734">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="1d373-735">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="1d373-735">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="1d373-736">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="1d373-736">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d373-737">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-737">Requirements</span></span>

|<span data-ttu-id="1d373-738">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-738">Requirement</span></span>| <span data-ttu-id="1d373-739">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-740">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-741">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-741">1.0</span></span>|
|[<span data-ttu-id="1d373-742">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-742">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-743">Restrito</span><span class="sxs-lookup"><span data-stu-id="1d373-743">Restricted</span></span>|
|[<span data-ttu-id="1d373-744">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-744">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-745">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-745">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d373-746">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1d373-746">Returns:</span></span>

<span data-ttu-id="1d373-747">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="1d373-747">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="1d373-748">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="1d373-748">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="1d373-749">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="1d373-749">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="1d373-750">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="1d373-750">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="1d373-751">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="1d373-751">Value of `entityType`</span></span> | <span data-ttu-id="1d373-752">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="1d373-752">Type of objects in returned array</span></span> | <span data-ttu-id="1d373-753">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="1d373-753">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="1d373-754">String</span><span class="sxs-lookup"><span data-stu-id="1d373-754">String</span></span> | <span data-ttu-id="1d373-755">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="1d373-755">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="1d373-756">Contato</span><span class="sxs-lookup"><span data-stu-id="1d373-756">Contact</span></span> | <span data-ttu-id="1d373-757">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1d373-757">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="1d373-758">String</span><span class="sxs-lookup"><span data-stu-id="1d373-758">String</span></span> | <span data-ttu-id="1d373-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1d373-759">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="1d373-760">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="1d373-760">MeetingSuggestion</span></span> | <span data-ttu-id="1d373-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1d373-761">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="1d373-762">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="1d373-762">PhoneNumber</span></span> | <span data-ttu-id="1d373-763">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="1d373-763">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="1d373-764">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="1d373-764">TaskSuggestion</span></span> | <span data-ttu-id="1d373-765">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1d373-765">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="1d373-766">String</span><span class="sxs-lookup"><span data-stu-id="1d373-766">String</span></span> | <span data-ttu-id="1d373-767">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="1d373-767">**Restricted**</span></span> |

<span data-ttu-id="1d373-768">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1d373-768">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="1d373-769">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-769">Example</span></span>

<span data-ttu-id="1d373-770">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1d373-770">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="1d373-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1d373-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1d373-772">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="1d373-772">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1d373-773">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d373-773">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d373-774">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="1d373-774">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d373-775">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="1d373-775">Parameters</span></span>

|<span data-ttu-id="1d373-776">Nome</span><span class="sxs-lookup"><span data-stu-id="1d373-776">Name</span></span>| <span data-ttu-id="1d373-777">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-777">Type</span></span>| <span data-ttu-id="1d373-778">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-778">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1d373-779">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d373-779">String</span></span>|<span data-ttu-id="1d373-780">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="1d373-780">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d373-781">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-781">Requirements</span></span>

|<span data-ttu-id="1d373-782">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-782">Requirement</span></span>| <span data-ttu-id="1d373-783">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-783">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-784">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-784">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-785">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-785">1.0</span></span>|
|[<span data-ttu-id="1d373-786">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-786">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-787">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-787">ReadItem</span></span>|
|[<span data-ttu-id="1d373-788">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-788">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-789">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-789">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d373-790">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1d373-790">Returns:</span></span>

<span data-ttu-id="1d373-p153">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="1d373-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="1d373-793">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1d373-793">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="1d373-794">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1d373-794">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="1d373-795">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="1d373-795">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1d373-796">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d373-796">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d373-p154">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="1d373-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1d373-800">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="1d373-800">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1d373-801">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="1d373-801">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="1d373-p155">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="1d373-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d373-804">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-804">Requirements</span></span>

|<span data-ttu-id="1d373-805">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-805">Requirement</span></span>| <span data-ttu-id="1d373-806">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-807">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-808">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-808">1.0</span></span>|
|[<span data-ttu-id="1d373-809">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-809">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-810">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-810">ReadItem</span></span>|
|[<span data-ttu-id="1d373-811">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-811">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-812">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-812">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d373-813">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1d373-813">Returns:</span></span>

<span data-ttu-id="1d373-p156">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="1d373-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="1d373-816">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="1d373-816">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1d373-817">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d373-817">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1d373-818">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-818">Example</span></span>

<span data-ttu-id="1d373-819">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="1d373-819">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="1d373-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="1d373-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="1d373-821">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="1d373-821">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1d373-822">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d373-822">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d373-823">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="1d373-823">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="1d373-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="1d373-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d373-826">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="1d373-826">Parameters</span></span>

|<span data-ttu-id="1d373-827">Nome</span><span class="sxs-lookup"><span data-stu-id="1d373-827">Name</span></span>| <span data-ttu-id="1d373-828">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-828">Type</span></span>| <span data-ttu-id="1d373-829">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-829">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1d373-830">String</span><span class="sxs-lookup"><span data-stu-id="1d373-830">String</span></span>|<span data-ttu-id="1d373-831">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="1d373-831">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d373-832">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-832">Requirements</span></span>

|<span data-ttu-id="1d373-833">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-833">Requirement</span></span>| <span data-ttu-id="1d373-834">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-835">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-836">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-836">1.0</span></span>|
|[<span data-ttu-id="1d373-837">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-838">ReadItem</span></span>|
|[<span data-ttu-id="1d373-839">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-840">Read</span><span class="sxs-lookup"><span data-stu-id="1d373-840">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d373-841">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1d373-841">Returns:</span></span>

<span data-ttu-id="1d373-842">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="1d373-842">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="1d373-843">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="1d373-843">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1d373-844">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="1d373-844">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1d373-845">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-845">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="1d373-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="1d373-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="1d373-847">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1d373-847">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="1d373-p158">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="1d373-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d373-850">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="1d373-850">Parameters</span></span>

|<span data-ttu-id="1d373-851">Nome</span><span class="sxs-lookup"><span data-stu-id="1d373-851">Name</span></span>| <span data-ttu-id="1d373-852">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-852">Type</span></span>| <span data-ttu-id="1d373-853">Atributos</span><span class="sxs-lookup"><span data-stu-id="1d373-853">Attributes</span></span>| <span data-ttu-id="1d373-854">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-854">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="1d373-855">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1d373-855">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="1d373-p159">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="1d373-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="1d373-859">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d373-859">Object</span></span>| <span data-ttu-id="1d373-860">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-860">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-861">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1d373-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1d373-862">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d373-862">Object</span></span>| <span data-ttu-id="1d373-863">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-863">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-864">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1d373-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1d373-865">function</span><span class="sxs-lookup"><span data-stu-id="1d373-865">function</span></span>||<span data-ttu-id="1d373-866">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d373-866">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1d373-867">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="1d373-867">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="1d373-868">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="1d373-868">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d373-869">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-869">Requirements</span></span>

|<span data-ttu-id="1d373-870">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-870">Requirement</span></span>| <span data-ttu-id="1d373-871">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-872">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-873">1.2</span><span class="sxs-lookup"><span data-stu-id="1d373-873">1.2</span></span>|
|[<span data-ttu-id="1d373-874">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-874">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-875">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1d373-875">ReadWriteItem</span></span>|
|[<span data-ttu-id="1d373-876">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-876">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-877">Escrever</span><span class="sxs-lookup"><span data-stu-id="1d373-877">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d373-878">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1d373-878">Returns:</span></span>

<span data-ttu-id="1d373-879">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="1d373-879">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="1d373-880">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="1d373-880">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1d373-881">String</span><span class="sxs-lookup"><span data-stu-id="1d373-881">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1d373-882">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-882">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="1d373-883">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1d373-883">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="1d373-884">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="1d373-884">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="1d373-p161">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="1d373-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d373-888">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="1d373-888">Parameters</span></span>

|<span data-ttu-id="1d373-889">Nome</span><span class="sxs-lookup"><span data-stu-id="1d373-889">Name</span></span>| <span data-ttu-id="1d373-890">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-890">Type</span></span>| <span data-ttu-id="1d373-891">Atributos</span><span class="sxs-lookup"><span data-stu-id="1d373-891">Attributes</span></span>| <span data-ttu-id="1d373-892">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-892">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1d373-893">function</span><span class="sxs-lookup"><span data-stu-id="1d373-893">function</span></span>||<span data-ttu-id="1d373-894">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d373-894">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1d373-895">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d373-895">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1d373-896">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="1d373-896">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="1d373-897">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d373-897">Object</span></span>| <span data-ttu-id="1d373-898">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-898">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-899">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1d373-899">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="1d373-900">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1d373-900">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d373-901">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-901">Requirements</span></span>

|<span data-ttu-id="1d373-902">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-902">Requirement</span></span>| <span data-ttu-id="1d373-903">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-904">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-905">1.0</span><span class="sxs-lookup"><span data-stu-id="1d373-905">1.0</span></span>|
|[<span data-ttu-id="1d373-906">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-906">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-907">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d373-907">ReadItem</span></span>|
|[<span data-ttu-id="1d373-908">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-908">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-909">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d373-909">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-910">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-910">Example</span></span>

<span data-ttu-id="1d373-p164">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="1d373-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="1d373-914">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1d373-914">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="1d373-915">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="1d373-915">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="1d373-p165">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="1d373-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d373-920">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="1d373-920">Parameters</span></span>

|<span data-ttu-id="1d373-921">Nome</span><span class="sxs-lookup"><span data-stu-id="1d373-921">Name</span></span>| <span data-ttu-id="1d373-922">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-922">Type</span></span>| <span data-ttu-id="1d373-923">Atributos</span><span class="sxs-lookup"><span data-stu-id="1d373-923">Attributes</span></span>| <span data-ttu-id="1d373-924">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-924">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="1d373-925">String</span><span class="sxs-lookup"><span data-stu-id="1d373-925">String</span></span>||<span data-ttu-id="1d373-926">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="1d373-926">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="1d373-927">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d373-927">Object</span></span>| <span data-ttu-id="1d373-928">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-928">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-929">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1d373-929">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1d373-930">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d373-930">Object</span></span>| <span data-ttu-id="1d373-931">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-931">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-932">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1d373-932">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1d373-933">function</span><span class="sxs-lookup"><span data-stu-id="1d373-933">function</span></span>| <span data-ttu-id="1d373-934">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-934">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-935">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d373-935">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1d373-936">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="1d373-936">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1d373-937">Erros</span><span class="sxs-lookup"><span data-stu-id="1d373-937">Errors</span></span>

| <span data-ttu-id="1d373-938">Código de erro</span><span class="sxs-lookup"><span data-stu-id="1d373-938">Error code</span></span> | <span data-ttu-id="1d373-939">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-939">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="1d373-940">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="1d373-940">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1d373-941">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-941">Requirements</span></span>

|<span data-ttu-id="1d373-942">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-942">Requirement</span></span>| <span data-ttu-id="1d373-943">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-944">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-945">1.1</span><span class="sxs-lookup"><span data-stu-id="1d373-945">1.1</span></span>|
|[<span data-ttu-id="1d373-946">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-946">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-947">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1d373-947">ReadWriteItem</span></span>|
|[<span data-ttu-id="1d373-948">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-948">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-949">Escrever</span><span class="sxs-lookup"><span data-stu-id="1d373-949">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-950">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-950">Example</span></span>

<span data-ttu-id="1d373-951">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="1d373-951">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="1d373-952">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="1d373-952">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="1d373-953">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1d373-953">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="1d373-p166">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="1d373-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d373-957">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="1d373-957">Parameters</span></span>

|<span data-ttu-id="1d373-958">Nome</span><span class="sxs-lookup"><span data-stu-id="1d373-958">Name</span></span>| <span data-ttu-id="1d373-959">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d373-959">Type</span></span>| <span data-ttu-id="1d373-960">Atributos</span><span class="sxs-lookup"><span data-stu-id="1d373-960">Attributes</span></span>| <span data-ttu-id="1d373-961">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d373-961">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="1d373-962">String</span><span class="sxs-lookup"><span data-stu-id="1d373-962">String</span></span>||<span data-ttu-id="1d373-p167">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="1d373-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="1d373-966">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d373-966">Object</span></span>| <span data-ttu-id="1d373-967">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-967">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-968">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1d373-968">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1d373-969">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d373-969">Object</span></span>| <span data-ttu-id="1d373-970">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-970">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-971">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1d373-971">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="1d373-972">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1d373-972">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="1d373-973">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d373-973">&lt;optional&gt;</span></span>|<span data-ttu-id="1d373-p168">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="1d373-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="1d373-p169">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="1d373-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="1d373-978">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="1d373-978">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="1d373-979">function</span><span class="sxs-lookup"><span data-stu-id="1d373-979">function</span></span>||<span data-ttu-id="1d373-980">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d373-980">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1d373-981">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d373-981">Requirements</span></span>

|<span data-ttu-id="1d373-982">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d373-982">Requirement</span></span>| <span data-ttu-id="1d373-983">Valor</span><span class="sxs-lookup"><span data-stu-id="1d373-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d373-984">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d373-984">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d373-985">1.2</span><span class="sxs-lookup"><span data-stu-id="1d373-985">1.2</span></span>|
|[<span data-ttu-id="1d373-986">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d373-986">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d373-987">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1d373-987">ReadWriteItem</span></span>|
|[<span data-ttu-id="1d373-988">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d373-988">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d373-989">Escrever</span><span class="sxs-lookup"><span data-stu-id="1d373-989">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1d373-990">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d373-990">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
