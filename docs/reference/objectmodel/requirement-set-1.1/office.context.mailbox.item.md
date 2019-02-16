---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,1
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 5a43029a64c63dec3d48136ffe0a9c3c76e18b6c
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068158"
---
# <a name="item"></a><span data-ttu-id="a9369-102">item</span><span class="sxs-lookup"><span data-stu-id="a9369-102">item</span></span>

### <span data-ttu-id="a9369-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="a9369-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="a9369-p102">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="a9369-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-107">Requirements</span></span>

|<span data-ttu-id="a9369-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-108">Requirement</span></span>| <span data-ttu-id="a9369-109">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-111">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-111">1.0</span></span>|
|[<span data-ttu-id="a9369-112">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-113">Restrito</span><span class="sxs-lookup"><span data-stu-id="a9369-113">Restricted</span></span>|
|[<span data-ttu-id="a9369-114">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-115">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="a9369-116">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-116">Example</span></span>

<span data-ttu-id="a9369-117">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="a9369-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="a9369-118">Membros</span><span class="sxs-lookup"><span data-stu-id="a9369-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="a9369-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a9369-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="a9369-p103">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a9369-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-122">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="a9369-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="a9369-123">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="a9369-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-124">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-124">Type</span></span>

*   <span data-ttu-id="a9369-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a9369-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-126">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-126">Requirements</span></span>

|<span data-ttu-id="a9369-127">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-127">Requirement</span></span>| <span data-ttu-id="a9369-128">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-129">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-130">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-130">1.0</span></span>|
|[<span data-ttu-id="a9369-131">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-132">ReadItem</span></span>|
|[<span data-ttu-id="a9369-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-134">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-135">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-135">Example</span></span>

<span data-ttu-id="a9369-136">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="a9369-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="a9369-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a9369-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="a9369-138">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a9369-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="a9369-139">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="a9369-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-140">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-140">Type</span></span>

*   [<span data-ttu-id="a9369-141">Destinatários</span><span class="sxs-lookup"><span data-stu-id="a9369-141">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="a9369-142">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-142">Requirements</span></span>

|<span data-ttu-id="a9369-143">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-143">Requirement</span></span>| <span data-ttu-id="a9369-144">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-145">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-146">1.1</span><span class="sxs-lookup"><span data-stu-id="a9369-146">1.1</span></span>|
|[<span data-ttu-id="a9369-147">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-148">ReadItem</span></span>|
|[<span data-ttu-id="a9369-149">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-150">Escrever</span><span class="sxs-lookup"><span data-stu-id="a9369-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-151">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="a9369-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="a9369-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="a9369-153">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="a9369-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-154">Type</span></span>

*   [<span data-ttu-id="a9369-155">Corpo</span><span class="sxs-lookup"><span data-stu-id="a9369-155">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="a9369-156">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-156">Requirements</span></span>

|<span data-ttu-id="a9369-157">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-157">Requirement</span></span>| <span data-ttu-id="a9369-158">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-159">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-160">1.1</span><span class="sxs-lookup"><span data-stu-id="a9369-160">1.1</span></span>|
|[<span data-ttu-id="a9369-161">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-162">ReadItem</span></span>|
|[<span data-ttu-id="a9369-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-164">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-165">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-165">Example</span></span>

<span data-ttu-id="a9369-166">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="a9369-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="a9369-167">A seguir apresentamos um exemplo do parâmetro result passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a9369-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="a9369-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a9369-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="a9369-169">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a9369-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="a9369-170">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a9369-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a9369-171">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-171">Read mode</span></span>

<span data-ttu-id="a9369-p107">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="a9369-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="a9369-174">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a9369-174">Compose mode</span></span>

<span data-ttu-id="a9369-175">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="a9369-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a9369-176">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-176">Type</span></span>

*   <span data-ttu-id="a9369-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a9369-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-178">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-178">Requirements</span></span>

|<span data-ttu-id="a9369-179">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-179">Requirement</span></span>| <span data-ttu-id="a9369-180">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-181">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-182">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-182">1.0</span></span>|
|[<span data-ttu-id="a9369-183">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-183">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-184">ReadItem</span></span>|
|[<span data-ttu-id="a9369-185">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-185">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-186">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-186">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="a9369-187">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="a9369-187">(nullable) conversationId :String</span></span>

<span data-ttu-id="a9369-188">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="a9369-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="a9369-p108">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="a9369-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="a9369-p109">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="a9369-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-193">Type</span></span>

*   <span data-ttu-id="a9369-194">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a9369-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-195">Requirements</span></span>

|<span data-ttu-id="a9369-196">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-196">Requirement</span></span>| <span data-ttu-id="a9369-197">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-199">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-199">1.0</span></span>|
|[<span data-ttu-id="a9369-200">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-200">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-201">ReadItem</span></span>|
|[<span data-ttu-id="a9369-202">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-202">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-203">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-204">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="a9369-205">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="a9369-205">dateTimeCreated :Date</span></span>

<span data-ttu-id="a9369-p110">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a9369-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-208">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-208">Type</span></span>

*   <span data-ttu-id="a9369-209">Data</span><span class="sxs-lookup"><span data-stu-id="a9369-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-210">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-210">Requirements</span></span>

|<span data-ttu-id="a9369-211">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-211">Requirement</span></span>| <span data-ttu-id="a9369-212">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-213">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-214">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-214">1.0</span></span>|
|[<span data-ttu-id="a9369-215">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-215">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-216">ReadItem</span></span>|
|[<span data-ttu-id="a9369-217">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-217">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-218">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-219">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="a9369-220">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="a9369-220">dateTimeModified :Date</span></span>

<span data-ttu-id="a9369-p111">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a9369-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-223">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a9369-223">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-224">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-224">Type</span></span>

*   <span data-ttu-id="a9369-225">Data</span><span class="sxs-lookup"><span data-stu-id="a9369-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-226">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-226">Requirements</span></span>

|<span data-ttu-id="a9369-227">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-227">Requirement</span></span>| <span data-ttu-id="a9369-228">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-229">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-230">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-230">1.0</span></span>|
|[<span data-ttu-id="a9369-231">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-231">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-232">ReadItem</span></span>|
|[<span data-ttu-id="a9369-233">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-233">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-234">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-235">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="a9369-236">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="a9369-236">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="a9369-237">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="a9369-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="a9369-p112">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="a9369-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a9369-240">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-240">Read mode</span></span>

<span data-ttu-id="a9369-241">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="a9369-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="a9369-242">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a9369-242">Compose mode</span></span>

<span data-ttu-id="a9369-243">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a9369-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="a9369-244">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="a9369-244">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="a9369-245">O exemplo a seguir define a hora de término de um compromisso usando [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) o método do `Time` objeto.</span><span class="sxs-lookup"><span data-stu-id="a9369-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="a9369-246">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-246">Type</span></span>

*   <span data-ttu-id="a9369-247">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="a9369-247">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-248">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-248">Requirements</span></span>

|<span data-ttu-id="a9369-249">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-249">Requirement</span></span>| <span data-ttu-id="a9369-250">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-251">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-252">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-252">1.0</span></span>|
|[<span data-ttu-id="a9369-253">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-253">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-254">ReadItem</span></span>|
|[<span data-ttu-id="a9369-255">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-255">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-256">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="a9369-257">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a9369-257">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="a9369-p113">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a9369-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="a9369-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="a9369-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-262">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a9369-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-263">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-263">Type</span></span>

*   [<span data-ttu-id="a9369-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a9369-264">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a9369-265">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-265">Requirements</span></span>

|<span data-ttu-id="a9369-266">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-266">Requirement</span></span>| <span data-ttu-id="a9369-267">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-268">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-269">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-269">1.0</span></span>|
|[<span data-ttu-id="a9369-270">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-270">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-271">ReadItem</span></span>|
|[<span data-ttu-id="a9369-272">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-272">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-273">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-274">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="a9369-275">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="a9369-275">internetMessageId :String</span></span>

<span data-ttu-id="a9369-p115">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a9369-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-278">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-278">Type</span></span>

*   <span data-ttu-id="a9369-279">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a9369-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-280">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-280">Requirements</span></span>

|<span data-ttu-id="a9369-281">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-281">Requirement</span></span>| <span data-ttu-id="a9369-282">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-283">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-284">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-284">1.0</span></span>|
|[<span data-ttu-id="a9369-285">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-285">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-286">ReadItem</span></span>|
|[<span data-ttu-id="a9369-287">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-287">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-288">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-289">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="a9369-290">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="a9369-290">itemClass :String</span></span>

<span data-ttu-id="a9369-p116">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a9369-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="a9369-p117">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="a9369-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="a9369-295">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-295">Type</span></span> | <span data-ttu-id="a9369-296">Descrição</span><span class="sxs-lookup"><span data-stu-id="a9369-296">Description</span></span> | <span data-ttu-id="a9369-297">classe de item</span><span class="sxs-lookup"><span data-stu-id="a9369-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="a9369-298">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="a9369-298">Appointment items</span></span> | <span data-ttu-id="a9369-299">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="a9369-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="a9369-300">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="a9369-300">Message items</span></span> | <span data-ttu-id="a9369-301">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="a9369-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="a9369-302">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="a9369-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-303">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-303">Type</span></span>

*   <span data-ttu-id="a9369-304">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a9369-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-305">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-305">Requirements</span></span>

|<span data-ttu-id="a9369-306">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-306">Requirement</span></span>| <span data-ttu-id="a9369-307">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-308">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-309">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-309">1.0</span></span>|
|[<span data-ttu-id="a9369-310">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-310">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-311">ReadItem</span></span>|
|[<span data-ttu-id="a9369-312">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-312">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-313">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-314">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="a9369-315">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="a9369-315">(nullable) itemId :String</span></span>

<span data-ttu-id="a9369-p118">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a9369-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-318">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="a9369-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a9369-319">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a9369-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="a9369-320">Antes de fazer chamadas API REST usando esse valor, ele deve ser convertido usando `Office.context.mailbox.convertToRestId`, que está disponível a partir do conjunto de requisitos 1.3.</span><span class="sxs-lookup"><span data-stu-id="a9369-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="a9369-321">Para saber mais, consulte [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="a9369-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-322">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-322">Type</span></span>

*   <span data-ttu-id="a9369-323">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a9369-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-324">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-324">Requirements</span></span>

|<span data-ttu-id="a9369-325">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-325">Requirement</span></span>| <span data-ttu-id="a9369-326">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-327">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-328">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-328">1.0</span></span>|
|[<span data-ttu-id="a9369-329">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-330">ReadItem</span></span>|
|[<span data-ttu-id="a9369-331">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-332">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-333">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-333">Example</span></span>

<span data-ttu-id="a9369-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="a9369-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="a9369-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="a9369-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="a9369-337">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="a9369-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="a9369-338">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a9369-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-339">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-339">Type</span></span>

*   [<span data-ttu-id="a9369-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="a9369-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="a9369-341">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-341">Requirements</span></span>

|<span data-ttu-id="a9369-342">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-342">Requirement</span></span>| <span data-ttu-id="a9369-343">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-344">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-345">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-345">1.0</span></span>|
|[<span data-ttu-id="a9369-346">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-346">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-347">ReadItem</span></span>|
|[<span data-ttu-id="a9369-348">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-348">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-349">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-350">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="a9369-351">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="a9369-351">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="a9369-352">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a9369-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a9369-353">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-353">Read mode</span></span>

<span data-ttu-id="a9369-354">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="a9369-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="a9369-355">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a9369-355">Compose mode</span></span>

<span data-ttu-id="a9369-356">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="a9369-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a9369-357">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-357">Type</span></span>

*   <span data-ttu-id="a9369-358">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="a9369-358">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-359">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-359">Requirements</span></span>

|<span data-ttu-id="a9369-360">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-360">Requirement</span></span>| <span data-ttu-id="a9369-361">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-362">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-363">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-363">1.0</span></span>|
|[<span data-ttu-id="a9369-364">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-365">ReadItem</span></span>|
|[<span data-ttu-id="a9369-366">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-367">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="a9369-368">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="a9369-368">normalizedSubject :String</span></span>

<span data-ttu-id="a9369-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a9369-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="a9369-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject).</span><span class="sxs-lookup"><span data-stu-id="a9369-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-373">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-373">Type</span></span>

*   <span data-ttu-id="a9369-374">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a9369-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-375">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-375">Requirements</span></span>

|<span data-ttu-id="a9369-376">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-376">Requirement</span></span>| <span data-ttu-id="a9369-377">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-378">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-379">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-379">1.0</span></span>|
|[<span data-ttu-id="a9369-380">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-380">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-381">ReadItem</span></span>|
|[<span data-ttu-id="a9369-382">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-382">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-383">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-384">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="a9369-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a9369-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="a9369-386">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="a9369-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="a9369-387">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a9369-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a9369-388">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-388">Read mode</span></span>

<span data-ttu-id="a9369-389">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="a9369-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="a9369-390">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a9369-390">Compose mode</span></span>

<span data-ttu-id="a9369-391">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="a9369-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a9369-392">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-392">Type</span></span>

*   <span data-ttu-id="a9369-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a9369-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-394">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-394">Requirements</span></span>

|<span data-ttu-id="a9369-395">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-395">Requirement</span></span>| <span data-ttu-id="a9369-396">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-397">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-398">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-398">1.0</span></span>|
|[<span data-ttu-id="a9369-399">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-399">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-400">ReadItem</span></span>|
|[<span data-ttu-id="a9369-401">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-401">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-402">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="a9369-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a9369-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="a9369-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a9369-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-406">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-406">Type</span></span>

*   [<span data-ttu-id="a9369-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a9369-407">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a9369-408">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-408">Requirements</span></span>

|<span data-ttu-id="a9369-409">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-409">Requirement</span></span>| <span data-ttu-id="a9369-410">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-411">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-412">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-412">1.0</span></span>|
|[<span data-ttu-id="a9369-413">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-414">ReadItem</span></span>|
|[<span data-ttu-id="a9369-415">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-416">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-417">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="a9369-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a9369-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="a9369-419">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="a9369-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="a9369-420">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a9369-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a9369-421">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-421">Read mode</span></span>

<span data-ttu-id="a9369-422">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="a9369-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="a9369-423">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a9369-423">Compose mode</span></span>

<span data-ttu-id="a9369-424">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="a9369-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="a9369-425">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-425">Type</span></span>

*   <span data-ttu-id="a9369-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a9369-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-427">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-427">Requirements</span></span>

|<span data-ttu-id="a9369-428">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-428">Requirement</span></span>| <span data-ttu-id="a9369-429">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-430">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-431">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-431">1.0</span></span>|
|[<span data-ttu-id="a9369-432">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-432">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-433">ReadItem</span></span>|
|[<span data-ttu-id="a9369-434">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-434">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-435">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="a9369-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a9369-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="a9369-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a9369-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="a9369-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="a9369-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-441">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a9369-441">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a9369-442">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-442">Type</span></span>

*   [<span data-ttu-id="a9369-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a9369-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a9369-444">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-444">Requirements</span></span>

|<span data-ttu-id="a9369-445">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-445">Requirement</span></span>| <span data-ttu-id="a9369-446">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-447">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-448">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-448">1.0</span></span>|
|[<span data-ttu-id="a9369-449">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-450">ReadItem</span></span>|
|[<span data-ttu-id="a9369-451">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-452">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-453">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="a9369-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="a9369-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="a9369-455">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="a9369-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="a9369-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="a9369-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a9369-458">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-458">Read mode</span></span>

<span data-ttu-id="a9369-459">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="a9369-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="a9369-460">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a9369-460">Compose mode</span></span>

<span data-ttu-id="a9369-461">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a9369-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="a9369-462">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="a9369-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="a9369-463">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a9369-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="a9369-464">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-464">Type</span></span>

*   <span data-ttu-id="a9369-465">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="a9369-465">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-466">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-466">Requirements</span></span>

|<span data-ttu-id="a9369-467">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-467">Requirement</span></span>| <span data-ttu-id="a9369-468">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-469">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-470">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-470">1.0</span></span>|
|[<span data-ttu-id="a9369-471">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-472">ReadItem</span></span>|
|[<span data-ttu-id="a9369-473">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-474">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-474">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="a9369-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a9369-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="a9369-476">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="a9369-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="a9369-477">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="a9369-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a9369-478">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-478">Read mode</span></span>

<span data-ttu-id="a9369-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="a9369-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="a9369-481">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a9369-481">Compose mode</span></span>

<span data-ttu-id="a9369-482">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="a9369-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="a9369-483">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-483">Type</span></span>

*   <span data-ttu-id="a9369-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a9369-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-485">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-485">Requirements</span></span>

|<span data-ttu-id="a9369-486">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-486">Requirement</span></span>| <span data-ttu-id="a9369-487">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-488">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-489">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-489">1.0</span></span>|
|[<span data-ttu-id="a9369-490">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-491">ReadItem</span></span>|
|[<span data-ttu-id="a9369-492">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-493">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-493">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="a9369-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a9369-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="a9369-495">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a9369-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="a9369-496">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a9369-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a9369-497">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-497">Read mode</span></span>

<span data-ttu-id="a9369-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="a9369-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="a9369-500">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a9369-500">Compose mode</span></span>

<span data-ttu-id="a9369-501">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="a9369-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a9369-502">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-502">Type</span></span>

*   <span data-ttu-id="a9369-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a9369-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-504">Requirements</span></span>

|<span data-ttu-id="a9369-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-505">Requirement</span></span>| <span data-ttu-id="a9369-506">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-507">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-508">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-508">1.0</span></span>|
|[<span data-ttu-id="a9369-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-510">ReadItem</span></span>|
|[<span data-ttu-id="a9369-511">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-512">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="a9369-513">Métodos</span><span class="sxs-lookup"><span data-stu-id="a9369-513">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="a9369-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a9369-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a9369-515">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="a9369-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a9369-516">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="a9369-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="a9369-517">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a9369-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a9369-518">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a9369-518">Parameters</span></span>

|<span data-ttu-id="a9369-519">Nome</span><span class="sxs-lookup"><span data-stu-id="a9369-519">Name</span></span>| <span data-ttu-id="a9369-520">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-520">Type</span></span>| <span data-ttu-id="a9369-521">Atributos</span><span class="sxs-lookup"><span data-stu-id="a9369-521">Attributes</span></span>| <span data-ttu-id="a9369-522">Descrição</span><span class="sxs-lookup"><span data-stu-id="a9369-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="a9369-523">String</span><span class="sxs-lookup"><span data-stu-id="a9369-523">String</span></span>||<span data-ttu-id="a9369-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a9369-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="a9369-526">String</span><span class="sxs-lookup"><span data-stu-id="a9369-526">String</span></span>||<span data-ttu-id="a9369-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a9369-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="a9369-529">Object</span><span class="sxs-lookup"><span data-stu-id="a9369-529">Object</span></span>| <span data-ttu-id="a9369-530">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-530">&lt;optional&gt;</span></span>|<span data-ttu-id="a9369-531">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a9369-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a9369-532">Objeto</span><span class="sxs-lookup"><span data-stu-id="a9369-532">Object</span></span>| <span data-ttu-id="a9369-533">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-533">&lt;optional&gt;</span></span>|<span data-ttu-id="a9369-534">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a9369-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a9369-535">function</span><span class="sxs-lookup"><span data-stu-id="a9369-535">function</span></span>| <span data-ttu-id="a9369-536">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-536">&lt;optional&gt;</span></span>|<span data-ttu-id="a9369-537">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a9369-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a9369-538">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a9369-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a9369-539">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="a9369-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a9369-540">Erros</span><span class="sxs-lookup"><span data-stu-id="a9369-540">Errors</span></span>

| <span data-ttu-id="a9369-541">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a9369-541">Error code</span></span> | <span data-ttu-id="a9369-542">Descrição</span><span class="sxs-lookup"><span data-stu-id="a9369-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="a9369-543">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="a9369-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="a9369-544">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="a9369-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="a9369-545">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="a9369-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a9369-546">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-546">Requirements</span></span>

|<span data-ttu-id="a9369-547">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-547">Requirement</span></span>| <span data-ttu-id="a9369-548">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-549">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-550">1.1</span><span class="sxs-lookup"><span data-stu-id="a9369-550">1.1</span></span>|
|[<span data-ttu-id="a9369-551">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-551">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a9369-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="a9369-553">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-554">Escrever</span><span class="sxs-lookup"><span data-stu-id="a9369-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-555">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-555">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="a9369-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a9369-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a9369-557">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="a9369-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="a9369-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="a9369-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="a9369-561">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a9369-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="a9369-562">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="a9369-562">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a9369-563">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a9369-563">Parameters</span></span>

|<span data-ttu-id="a9369-564">Nome</span><span class="sxs-lookup"><span data-stu-id="a9369-564">Name</span></span>| <span data-ttu-id="a9369-565">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-565">Type</span></span>| <span data-ttu-id="a9369-566">Atributos</span><span class="sxs-lookup"><span data-stu-id="a9369-566">Attributes</span></span>| <span data-ttu-id="a9369-567">Descrição</span><span class="sxs-lookup"><span data-stu-id="a9369-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="a9369-568">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a9369-568">String</span></span>||<span data-ttu-id="a9369-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a9369-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="a9369-571">String</span><span class="sxs-lookup"><span data-stu-id="a9369-571">String</span></span>||<span data-ttu-id="a9369-572">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="a9369-572">The subject of the item to be attached.</span></span> <span data-ttu-id="a9369-573">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a9369-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="a9369-574">Object</span><span class="sxs-lookup"><span data-stu-id="a9369-574">Object</span></span>| <span data-ttu-id="a9369-575">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-575">&lt;optional&gt;</span></span>|<span data-ttu-id="a9369-576">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a9369-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a9369-577">Objeto</span><span class="sxs-lookup"><span data-stu-id="a9369-577">Object</span></span>| <span data-ttu-id="a9369-578">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-578">&lt;optional&gt;</span></span>|<span data-ttu-id="a9369-579">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a9369-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a9369-580">function</span><span class="sxs-lookup"><span data-stu-id="a9369-580">function</span></span>| <span data-ttu-id="a9369-581">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-581">&lt;optional&gt;</span></span>|<span data-ttu-id="a9369-582">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a9369-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a9369-583">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a9369-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a9369-584">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="a9369-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a9369-585">Erros</span><span class="sxs-lookup"><span data-stu-id="a9369-585">Errors</span></span>

| <span data-ttu-id="a9369-586">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a9369-586">Error code</span></span> | <span data-ttu-id="a9369-587">Descrição</span><span class="sxs-lookup"><span data-stu-id="a9369-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="a9369-588">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="a9369-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a9369-589">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-589">Requirements</span></span>

|<span data-ttu-id="a9369-590">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-590">Requirement</span></span>| <span data-ttu-id="a9369-591">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-592">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-593">1.1</span><span class="sxs-lookup"><span data-stu-id="a9369-593">1.1</span></span>|
|[<span data-ttu-id="a9369-594">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-594">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a9369-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="a9369-596">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-596">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-597">Escrever</span><span class="sxs-lookup"><span data-stu-id="a9369-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-598">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-598">Example</span></span>

<span data-ttu-id="a9369-599">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="a9369-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="a9369-600">displayReplyAllForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="a9369-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="a9369-601">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="a9369-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-602">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a9369-602">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a9369-603">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="a9369-603">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a9369-604">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="a9369-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-605">A capacidade de incluir anexos na chamada para `displayReplyAllForm` não tem suporte no conjunto de requisitos 1.1.</span><span class="sxs-lookup"><span data-stu-id="a9369-605">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="a9369-606">O suporte a anexos foi adicionado a `displayReplyAllForm` no conjunto de requisitos 1.2 e acima.</span><span class="sxs-lookup"><span data-stu-id="a9369-606">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a9369-607">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a9369-607">Parameters</span></span>

|<span data-ttu-id="a9369-608">Nome</span><span class="sxs-lookup"><span data-stu-id="a9369-608">Name</span></span>| <span data-ttu-id="a9369-609">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-609">Type</span></span>| <span data-ttu-id="a9369-610">Descrição</span><span class="sxs-lookup"><span data-stu-id="a9369-610">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="a9369-611">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a9369-611">String &#124; Object</span></span>| |<span data-ttu-id="a9369-p138">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a9369-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a9369-614">**OU**</span><span class="sxs-lookup"><span data-stu-id="a9369-614">**OR**</span></span><br/><span data-ttu-id="a9369-p139">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="a9369-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="a9369-617">String</span><span class="sxs-lookup"><span data-stu-id="a9369-617">String</span></span> | <span data-ttu-id="a9369-618">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-618">&lt;optional&gt;</span></span> | <span data-ttu-id="a9369-p140">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a9369-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="a9369-621">function</span><span class="sxs-lookup"><span data-stu-id="a9369-621">function</span></span> | <span data-ttu-id="a9369-622">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-622">&lt;optional&gt;</span></span> | <span data-ttu-id="a9369-623">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a9369-623">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a9369-624">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-624">Requirements</span></span>

|<span data-ttu-id="a9369-625">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-625">Requirement</span></span>| <span data-ttu-id="a9369-626">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-627">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-628">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-628">1.0</span></span>|
|[<span data-ttu-id="a9369-629">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-629">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-630">ReadItem</span></span>|
|[<span data-ttu-id="a9369-631">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-631">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-632">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-632">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a9369-633">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a9369-633">Examples</span></span>

<span data-ttu-id="a9369-634">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="a9369-634">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="a9369-635">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="a9369-635">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="a9369-636">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="a9369-636">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a9369-637">Responder com um corpo e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a9369-637">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="a9369-638">displayReplyForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="a9369-638">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="a9369-639">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="a9369-639">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-640">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a9369-640">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a9369-641">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="a9369-641">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a9369-642">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="a9369-642">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-643">A capacidade de incluir anexos na chamada para `displayReplyForm` não tem suporte no conjunto de requisitos 1.1.</span><span class="sxs-lookup"><span data-stu-id="a9369-643">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="a9369-644">O suporte a anexos foi adicionado a `displayReplyForm` no conjunto de requisitos 1.2 e acima.</span><span class="sxs-lookup"><span data-stu-id="a9369-644">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a9369-645">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a9369-645">Parameters</span></span>

|<span data-ttu-id="a9369-646">Nome</span><span class="sxs-lookup"><span data-stu-id="a9369-646">Name</span></span>| <span data-ttu-id="a9369-647">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-647">Type</span></span>| <span data-ttu-id="a9369-648">Descrição</span><span class="sxs-lookup"><span data-stu-id="a9369-648">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="a9369-649">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a9369-649">String &#124; Object</span></span>| | <span data-ttu-id="a9369-p142">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a9369-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a9369-652">**OU**</span><span class="sxs-lookup"><span data-stu-id="a9369-652">**OR**</span></span><br/><span data-ttu-id="a9369-p143">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="a9369-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="a9369-655">String</span><span class="sxs-lookup"><span data-stu-id="a9369-655">String</span></span> | <span data-ttu-id="a9369-656">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-656">&lt;optional&gt;</span></span> | <span data-ttu-id="a9369-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a9369-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="a9369-659">function</span><span class="sxs-lookup"><span data-stu-id="a9369-659">function</span></span> | <span data-ttu-id="a9369-660">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-660">&lt;optional&gt;</span></span> | <span data-ttu-id="a9369-661">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a9369-661">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a9369-662">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-662">Requirements</span></span>

|<span data-ttu-id="a9369-663">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-663">Requirement</span></span>| <span data-ttu-id="a9369-664">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-665">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-666">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-666">1.0</span></span>|
|[<span data-ttu-id="a9369-667">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-667">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-668">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-668">ReadItem</span></span>|
|[<span data-ttu-id="a9369-669">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-669">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-670">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-670">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a9369-671">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a9369-671">Examples</span></span>

<span data-ttu-id="a9369-672">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="a9369-672">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="a9369-673">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="a9369-673">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="a9369-674">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="a9369-674">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a9369-675">Responder com um corpo e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a9369-675">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="a9369-676">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a9369-676">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="a9369-677">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="a9369-677">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-678">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a9369-678">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-679">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-679">Requirements</span></span>

|<span data-ttu-id="a9369-680">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-680">Requirement</span></span>| <span data-ttu-id="a9369-681">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-681">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-682">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-682">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-683">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-683">1.0</span></span>|
|[<span data-ttu-id="a9369-684">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-684">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-685">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-685">ReadItem</span></span>|
|[<span data-ttu-id="a9369-686">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-686">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-687">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-687">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a9369-688">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a9369-688">Returns:</span></span>

<span data-ttu-id="a9369-689">Tipo: [Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a9369-689">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a9369-690">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-690">Example</span></span>

<span data-ttu-id="a9369-691">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a9369-691">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="a9369-692">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a9369-692">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a9369-693">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="a9369-693">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-694">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a9369-694">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a9369-695">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a9369-695">Parameters</span></span>

|<span data-ttu-id="a9369-696">Nome</span><span class="sxs-lookup"><span data-stu-id="a9369-696">Name</span></span>| <span data-ttu-id="a9369-697">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-697">Type</span></span>| <span data-ttu-id="a9369-698">Descrição</span><span class="sxs-lookup"><span data-stu-id="a9369-698">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="a9369-699">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="a9369-699">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="a9369-700">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="a9369-700">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a9369-701">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-701">Requirements</span></span>

|<span data-ttu-id="a9369-702">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-702">Requirement</span></span>| <span data-ttu-id="a9369-703">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-703">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-704">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-704">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-705">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-705">1.0</span></span>|
|[<span data-ttu-id="a9369-706">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-706">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-707">Restrito</span><span class="sxs-lookup"><span data-stu-id="a9369-707">Restricted</span></span>|
|[<span data-ttu-id="a9369-708">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-708">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-709">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-709">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a9369-710">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a9369-710">Returns:</span></span>

<span data-ttu-id="a9369-711">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="a9369-711">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="a9369-712">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="a9369-712">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="a9369-713">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="a9369-713">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="a9369-714">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="a9369-714">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="a9369-715">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="a9369-715">Value of `entityType`</span></span> | <span data-ttu-id="a9369-716">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="a9369-716">Type of objects in returned array</span></span> | <span data-ttu-id="a9369-717">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="a9369-717">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="a9369-718">String</span><span class="sxs-lookup"><span data-stu-id="a9369-718">String</span></span> | <span data-ttu-id="a9369-719">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="a9369-719">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="a9369-720">Contato</span><span class="sxs-lookup"><span data-stu-id="a9369-720">Contact</span></span> | <span data-ttu-id="a9369-721">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a9369-721">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="a9369-722">String</span><span class="sxs-lookup"><span data-stu-id="a9369-722">String</span></span> | <span data-ttu-id="a9369-723">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a9369-723">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="a9369-724">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="a9369-724">MeetingSuggestion</span></span> | <span data-ttu-id="a9369-725">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a9369-725">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="a9369-726">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="a9369-726">PhoneNumber</span></span> | <span data-ttu-id="a9369-727">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="a9369-727">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="a9369-728">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="a9369-728">TaskSuggestion</span></span> | <span data-ttu-id="a9369-729">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a9369-729">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="a9369-730">String</span><span class="sxs-lookup"><span data-stu-id="a9369-730">String</span></span> | <span data-ttu-id="a9369-731">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="a9369-731">**Restricted**</span></span> |

<span data-ttu-id="a9369-732">Tipo:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a9369-732">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="a9369-733">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-733">Example</span></span>

<span data-ttu-id="a9369-734">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a9369-734">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="a9369-735">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a9369-735">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a9369-736">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a9369-736">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-737">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a9369-737">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a9369-738">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="a9369-738">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a9369-739">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a9369-739">Parameters</span></span>

|<span data-ttu-id="a9369-740">Nome</span><span class="sxs-lookup"><span data-stu-id="a9369-740">Name</span></span>| <span data-ttu-id="a9369-741">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-741">Type</span></span>| <span data-ttu-id="a9369-742">Descrição</span><span class="sxs-lookup"><span data-stu-id="a9369-742">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="a9369-743">String</span><span class="sxs-lookup"><span data-stu-id="a9369-743">String</span></span>|<span data-ttu-id="a9369-744">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="a9369-744">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a9369-745">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-745">Requirements</span></span>

|<span data-ttu-id="a9369-746">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-746">Requirement</span></span>| <span data-ttu-id="a9369-747">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-748">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-749">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-749">1.0</span></span>|
|[<span data-ttu-id="a9369-750">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-751">ReadItem</span></span>|
|[<span data-ttu-id="a9369-752">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-753">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a9369-754">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a9369-754">Returns:</span></span>

<span data-ttu-id="a9369-p146">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="a9369-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="a9369-757">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a9369-757">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="a9369-758">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a9369-758">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="a9369-759">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a9369-759">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-760">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a9369-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a9369-p147">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="a9369-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a9369-764">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="a9369-764">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a9369-765">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a9369-765">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="a9369-p148">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="a9369-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9369-768">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-768">Requirements</span></span>

|<span data-ttu-id="a9369-769">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-769">Requirement</span></span>| <span data-ttu-id="a9369-770">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-770">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-771">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-771">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-772">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-772">1.0</span></span>|
|[<span data-ttu-id="a9369-773">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-773">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-774">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-774">ReadItem</span></span>|
|[<span data-ttu-id="a9369-775">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-775">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-776">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-776">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a9369-777">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a9369-777">Returns:</span></span>

<span data-ttu-id="a9369-p149">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="a9369-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="a9369-780">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="a9369-780">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a9369-781">Objeto</span><span class="sxs-lookup"><span data-stu-id="a9369-781">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a9369-782">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-782">Example</span></span>

<span data-ttu-id="a9369-783">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="a9369-783">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="a9369-784">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="a9369-784">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="a9369-785">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a9369-785">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a9369-786">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a9369-786">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a9369-787">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="a9369-787">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="a9369-p150">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="a9369-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a9369-790">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a9369-790">Parameters</span></span>

|<span data-ttu-id="a9369-791">Nome</span><span class="sxs-lookup"><span data-stu-id="a9369-791">Name</span></span>| <span data-ttu-id="a9369-792">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-792">Type</span></span>| <span data-ttu-id="a9369-793">Descrição</span><span class="sxs-lookup"><span data-stu-id="a9369-793">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="a9369-794">String</span><span class="sxs-lookup"><span data-stu-id="a9369-794">String</span></span>|<span data-ttu-id="a9369-795">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="a9369-795">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a9369-796">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-796">Requirements</span></span>

|<span data-ttu-id="a9369-797">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-797">Requirement</span></span>| <span data-ttu-id="a9369-798">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-798">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-799">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-799">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-800">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-800">1.0</span></span>|
|[<span data-ttu-id="a9369-801">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-801">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-802">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-802">ReadItem</span></span>|
|[<span data-ttu-id="a9369-803">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-803">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-804">Read</span><span class="sxs-lookup"><span data-stu-id="a9369-804">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a9369-805">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a9369-805">Returns:</span></span>

<span data-ttu-id="a9369-806">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a9369-806">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="a9369-807">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="a9369-807">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a9369-808">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="a9369-808">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a9369-809">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-809">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="a9369-810">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a9369-810">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="a9369-811">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="a9369-811">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="a9369-p151">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="a9369-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a9369-815">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a9369-815">Parameters</span></span>

|<span data-ttu-id="a9369-816">Nome</span><span class="sxs-lookup"><span data-stu-id="a9369-816">Name</span></span>| <span data-ttu-id="a9369-817">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-817">Type</span></span>| <span data-ttu-id="a9369-818">Atributos</span><span class="sxs-lookup"><span data-stu-id="a9369-818">Attributes</span></span>| <span data-ttu-id="a9369-819">Descrição</span><span class="sxs-lookup"><span data-stu-id="a9369-819">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="a9369-820">function</span><span class="sxs-lookup"><span data-stu-id="a9369-820">function</span></span>||<span data-ttu-id="a9369-821">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a9369-821">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a9369-822">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a9369-822">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a9369-823">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="a9369-823">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="a9369-824">Objeto</span><span class="sxs-lookup"><span data-stu-id="a9369-824">Object</span></span>| <span data-ttu-id="a9369-825">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-825">&lt;optional&gt;</span></span>|<span data-ttu-id="a9369-826">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a9369-826">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="a9369-827">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a9369-827">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a9369-828">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-828">Requirements</span></span>

|<span data-ttu-id="a9369-829">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-829">Requirement</span></span>| <span data-ttu-id="a9369-830">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-830">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-831">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-831">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-832">1.0</span><span class="sxs-lookup"><span data-stu-id="a9369-832">1.0</span></span>|
|[<span data-ttu-id="a9369-833">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-833">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-834">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a9369-834">ReadItem</span></span>|
|[<span data-ttu-id="a9369-835">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-835">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-836">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a9369-836">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-837">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-837">Example</span></span>

<span data-ttu-id="a9369-p154">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="a9369-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="a9369-841">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a9369-841">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="a9369-842">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a9369-842">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="a9369-p155">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="a9369-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a9369-847">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a9369-847">Parameters</span></span>

|<span data-ttu-id="a9369-848">Nome</span><span class="sxs-lookup"><span data-stu-id="a9369-848">Name</span></span>| <span data-ttu-id="a9369-849">Tipo</span><span class="sxs-lookup"><span data-stu-id="a9369-849">Type</span></span>| <span data-ttu-id="a9369-850">Atributos</span><span class="sxs-lookup"><span data-stu-id="a9369-850">Attributes</span></span>| <span data-ttu-id="a9369-851">Descrição</span><span class="sxs-lookup"><span data-stu-id="a9369-851">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="a9369-852">String</span><span class="sxs-lookup"><span data-stu-id="a9369-852">String</span></span>||<span data-ttu-id="a9369-853">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="a9369-853">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="a9369-854">Objeto</span><span class="sxs-lookup"><span data-stu-id="a9369-854">Object</span></span>| <span data-ttu-id="a9369-855">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-855">&lt;optional&gt;</span></span>|<span data-ttu-id="a9369-856">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a9369-856">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a9369-857">Objeto</span><span class="sxs-lookup"><span data-stu-id="a9369-857">Object</span></span>| <span data-ttu-id="a9369-858">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-858">&lt;optional&gt;</span></span>|<span data-ttu-id="a9369-859">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a9369-859">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a9369-860">function</span><span class="sxs-lookup"><span data-stu-id="a9369-860">function</span></span>| <span data-ttu-id="a9369-861">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a9369-861">&lt;optional&gt;</span></span>|<span data-ttu-id="a9369-862">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a9369-862">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a9369-863">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="a9369-863">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a9369-864">Erros</span><span class="sxs-lookup"><span data-stu-id="a9369-864">Errors</span></span>

| <span data-ttu-id="a9369-865">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a9369-865">Error code</span></span> | <span data-ttu-id="a9369-866">Descrição</span><span class="sxs-lookup"><span data-stu-id="a9369-866">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="a9369-867">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="a9369-867">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a9369-868">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a9369-868">Requirements</span></span>

|<span data-ttu-id="a9369-869">Requisito</span><span class="sxs-lookup"><span data-stu-id="a9369-869">Requirement</span></span>| <span data-ttu-id="a9369-870">Valor</span><span class="sxs-lookup"><span data-stu-id="a9369-870">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9369-871">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a9369-871">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9369-872">1.1</span><span class="sxs-lookup"><span data-stu-id="a9369-872">1.1</span></span>|
|[<span data-ttu-id="a9369-873">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a9369-873">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a9369-874">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a9369-874">ReadWriteItem</span></span>|
|[<span data-ttu-id="a9369-875">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a9369-875">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9369-876">Escrever</span><span class="sxs-lookup"><span data-stu-id="a9369-876">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a9369-877">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a9369-877">Example</span></span>

<span data-ttu-id="a9369-878">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="a9369-878">The following code removes an attachment with an identifier of '0'.</span></span>

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
