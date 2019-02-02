---
title: Office.Context.Mailbox.item - requisito definir 1,4
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: 711d9659430c4a904b1aad81991d5371ced3f282
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701880"
---
# <a name="item"></a><span data-ttu-id="f9c3e-102">item</span><span class="sxs-lookup"><span data-stu-id="f9c3e-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="f9c3e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f9c3e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="f9c3e-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-106">Requirements</span></span>

|<span data-ttu-id="f9c3e-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-107">Requirement</span></span>| <span data-ttu-id="f9c3e-108">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-110">1.0</span></span>|
|[<span data-ttu-id="f9c3e-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-112">Restricted</span></span>|
|[<span data-ttu-id="f9c3e-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-114">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="f9c3e-115">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-115">Example</span></span>

<span data-ttu-id="f9c3e-116">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
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
}
```

### <a name="members"></a><span data-ttu-id="f9c3e-117">Membros</span><span class="sxs-lookup"><span data-stu-id="f9c3e-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="f9c3e-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f9c3e-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="f9c3e-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-121">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f9c3e-122">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-123">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-123">Type:</span></span>

*   <span data-ttu-id="f9c3e-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f9c3e-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-125">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-125">Requirements</span></span>

|<span data-ttu-id="f9c3e-126">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-126">Requirement</span></span>| <span data-ttu-id="f9c3e-127">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-128">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-129">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-129">1.0</span></span>|
|[<span data-ttu-id="f9c3e-130">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-131">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-132">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-133">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-134">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-134">Example</span></span>

<span data-ttu-id="f9c3e-135">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f9c3e-136">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-136">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f9c3e-137">Obtém um objeto que fornece métodos para acessar ou atualizar a linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-137">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f9c3e-138">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-139">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-139">Type:</span></span>

*   [<span data-ttu-id="f9c3e-140">Destinatários</span><span class="sxs-lookup"><span data-stu-id="f9c3e-140">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="f9c3e-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-141">Requirements</span></span>

|<span data-ttu-id="f9c3e-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-142">Requirement</span></span>| <span data-ttu-id="f9c3e-143">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-145">1.1</span><span class="sxs-lookup"><span data-stu-id="f9c3e-145">1.1</span></span>|
|[<span data-ttu-id="f9c3e-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-147">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-149">Escrever</span><span class="sxs-lookup"><span data-stu-id="f9c3e-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-150">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="f9c3e-151">body :[Body](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-151">body :[Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="f9c3e-152">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-153">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-153">Type:</span></span>

*   [<span data-ttu-id="f9c3e-154">Corpo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-154">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="f9c3e-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-155">Requirements</span></span>

|<span data-ttu-id="f9c3e-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-156">Requirement</span></span>| <span data-ttu-id="f9c3e-157">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-159">1.1</span><span class="sxs-lookup"><span data-stu-id="f9c3e-159">1.1</span></span>|
|[<span data-ttu-id="f9c3e-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-161">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-163">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f9c3e-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f9c3e-165">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f9c3e-166">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f9c3e-167">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-167">Read mode</span></span>

<span data-ttu-id="f9c3e-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f9c3e-170">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f9c3e-170">Compose mode</span></span>

<span data-ttu-id="f9c3e-171">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-172">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-172">Type:</span></span>

*   <span data-ttu-id="f9c3e-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-174">Requirements</span></span>

|<span data-ttu-id="f9c3e-175">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-175">Requirement</span></span>| <span data-ttu-id="f9c3e-176">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-177">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-178">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-178">1.0</span></span>|
|[<span data-ttu-id="f9c3e-179">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-180">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-181">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-182">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-183">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-183">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="f9c3e-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="f9c3e-185">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f9c3e-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f9c3e-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-190">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-190">Type:</span></span>

*   <span data-ttu-id="f9c3e-191">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9c3e-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-192">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-192">Requirements</span></span>

|<span data-ttu-id="f9c3e-193">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-193">Requirement</span></span>| <span data-ttu-id="f9c3e-194">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-195">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-196">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-196">1.0</span></span>|
|[<span data-ttu-id="f9c3e-197">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-198">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-200">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="f9c3e-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="f9c3e-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="f9c3e-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-204">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-204">Type:</span></span>

*   <span data-ttu-id="f9c3e-205">Data</span><span class="sxs-lookup"><span data-stu-id="f9c3e-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-206">Requirements</span></span>

|<span data-ttu-id="f9c3e-207">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-207">Requirement</span></span>| <span data-ttu-id="f9c3e-208">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-209">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-210">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-210">1.0</span></span>|
|[<span data-ttu-id="f9c3e-211">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-212">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-213">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-214">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-215">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-215">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="f9c3e-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="f9c3e-216">dateTimeModified :Date</span></span>

<span data-ttu-id="f9c3e-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-219">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-220">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-220">Type:</span></span>

*   <span data-ttu-id="f9c3e-221">Data</span><span class="sxs-lookup"><span data-stu-id="f9c3e-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-222">Requirements</span></span>

|<span data-ttu-id="f9c3e-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-223">Requirement</span></span>| <span data-ttu-id="f9c3e-224">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-225">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-226">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-226">1.0</span></span>|
|[<span data-ttu-id="f9c3e-227">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-228">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-229">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-230">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-231">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-231">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="f9c3e-232">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-232">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="f9c3e-233">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f9c3e-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f9c3e-236">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-236">Read mode</span></span>

<span data-ttu-id="f9c3e-237">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f9c3e-238">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f9c3e-238">Compose mode</span></span>

<span data-ttu-id="f9c3e-239">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f9c3e-240">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-241">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-241">Type:</span></span>

*   <span data-ttu-id="f9c3e-242">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-242">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-243">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-243">Requirements</span></span>

|<span data-ttu-id="f9c3e-244">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-244">Requirement</span></span>| <span data-ttu-id="f9c3e-245">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-246">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-247">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-247">1.0</span></span>|
|[<span data-ttu-id="f9c3e-248">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-249">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-250">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-251">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-252">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-252">Example</span></span>

<span data-ttu-id="f9c3e-253">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="f9c3e-254">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-254">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="f9c3e-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="f9c3e-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-259">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-260">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-260">Type:</span></span>

*   [<span data-ttu-id="f9c3e-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f9c3e-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f9c3e-262">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-262">Requirements</span></span>

|<span data-ttu-id="f9c3e-263">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-263">Requirement</span></span>| <span data-ttu-id="f9c3e-264">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-265">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-266">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-266">1.0</span></span>|
|[<span data-ttu-id="f9c3e-267">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-268">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-269">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-270">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="f9c3e-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-271">internetMessageId :String</span></span>

<span data-ttu-id="f9c3e-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-274">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-274">Type:</span></span>

*   <span data-ttu-id="f9c3e-275">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9c3e-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-276">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-276">Requirements</span></span>

|<span data-ttu-id="f9c3e-277">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-277">Requirement</span></span>| <span data-ttu-id="f9c3e-278">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-279">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-280">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-280">1.0</span></span>|
|[<span data-ttu-id="f9c3e-281">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-282">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-283">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-284">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-285">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-285">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="f9c3e-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-286">itemClass :String</span></span>

<span data-ttu-id="f9c3e-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f9c3e-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="f9c3e-291">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-291">Type</span></span> | <span data-ttu-id="f9c3e-292">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-292">Description</span></span> | <span data-ttu-id="f9c3e-293">classe de item</span><span class="sxs-lookup"><span data-stu-id="f9c3e-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="f9c3e-294">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="f9c3e-294">Appointment items</span></span> | <span data-ttu-id="f9c3e-295">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="f9c3e-296">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-296">Message items</span></span> | <span data-ttu-id="f9c3e-297">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="f9c3e-298">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-299">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-299">Type:</span></span>

*   <span data-ttu-id="f9c3e-300">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9c3e-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-301">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-301">Requirements</span></span>

|<span data-ttu-id="f9c3e-302">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-302">Requirement</span></span>| <span data-ttu-id="f9c3e-303">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-304">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-305">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-305">1.0</span></span>|
|[<span data-ttu-id="f9c3e-306">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-307">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-308">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-309">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-310">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-310">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f9c3e-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-311">(nullable) itemId :String</span></span>

<span data-ttu-id="f9c3e-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-314">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f9c3e-315">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f9c3e-316">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-316">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f9c3e-317">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="f9c3e-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-320">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-320">Type:</span></span>

*   <span data-ttu-id="f9c3e-321">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9c3e-321">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-322">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-322">Requirements</span></span>

|<span data-ttu-id="f9c3e-323">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-323">Requirement</span></span>| <span data-ttu-id="f9c3e-324">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-324">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-325">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-325">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-326">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-326">1.0</span></span>|
|[<span data-ttu-id="f9c3e-327">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-327">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-328">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-328">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-329">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-329">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-330">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-330">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-331">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-331">Example</span></span>

<span data-ttu-id="f9c3e-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="f9c3e-334">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-334">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="f9c3e-335">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-335">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f9c3e-336">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-336">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-337">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-337">Type:</span></span>

*   [<span data-ttu-id="f9c3e-338">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="f9c3e-338">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="f9c3e-339">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-339">Requirements</span></span>

|<span data-ttu-id="f9c3e-340">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-340">Requirement</span></span>| <span data-ttu-id="f9c3e-341">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-342">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-343">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-343">1.0</span></span>|
|[<span data-ttu-id="f9c3e-344">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-345">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-346">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-347">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-348">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-348">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="f9c3e-349">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-349">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="f9c3e-350">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-350">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f9c3e-351">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-351">Read mode</span></span>

<span data-ttu-id="f9c3e-352">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-352">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f9c3e-353">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f9c3e-353">Compose mode</span></span>

<span data-ttu-id="f9c3e-354">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-354">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-355">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-355">Type:</span></span>

*   <span data-ttu-id="f9c3e-356">String | [Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-356">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-357">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-357">Requirements</span></span>

|<span data-ttu-id="f9c3e-358">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-358">Requirement</span></span>| <span data-ttu-id="f9c3e-359">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-360">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-361">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-361">1.0</span></span>|
|[<span data-ttu-id="f9c3e-362">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-363">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-364">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-365">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-366">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-366">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f9c3e-367">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-367">normalizedSubject :String</span></span>

<span data-ttu-id="f9c3e-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f9c3e-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-372">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-372">Type:</span></span>

*   <span data-ttu-id="f9c3e-373">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9c3e-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-374">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-374">Requirements</span></span>

|<span data-ttu-id="f9c3e-375">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-375">Requirement</span></span>| <span data-ttu-id="f9c3e-376">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-377">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-378">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-378">1.0</span></span>|
|[<span data-ttu-id="f9c3e-379">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-379">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-380">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-381">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-381">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-382">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-383">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-383">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="f9c3e-384">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-384">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="f9c3e-385">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-385">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-386">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-386">Type:</span></span>

*   [<span data-ttu-id="f9c3e-387">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="f9c3e-387">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="f9c3e-388">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-388">Requirements</span></span>

|<span data-ttu-id="f9c3e-389">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-389">Requirement</span></span>| <span data-ttu-id="f9c3e-390">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-390">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-391">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-392">1.3</span><span class="sxs-lookup"><span data-stu-id="f9c3e-392">1.3</span></span>|
|[<span data-ttu-id="f9c3e-393">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-393">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-394">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-394">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-395">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-395">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-396">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-396">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f9c3e-397">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-397">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f9c3e-398">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-398">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f9c3e-399">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-399">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f9c3e-400">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-400">Read mode</span></span>

<span data-ttu-id="f9c3e-401">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-401">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f9c3e-402">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f9c3e-402">Compose mode</span></span>

<span data-ttu-id="f9c3e-403">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-403">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-404">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-404">Type:</span></span>

*   <span data-ttu-id="f9c3e-405">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-405">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-406">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-406">Requirements</span></span>

|<span data-ttu-id="f9c3e-407">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-407">Requirement</span></span>| <span data-ttu-id="f9c3e-408">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-409">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-410">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-410">1.0</span></span>|
|[<span data-ttu-id="f9c3e-411">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-412">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-413">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-414">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-414">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-415">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-415">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="f9c3e-416">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-416">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="f9c3e-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-419">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-419">Type:</span></span>

*   [<span data-ttu-id="f9c3e-420">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f9c3e-420">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f9c3e-421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-421">Requirements</span></span>

|<span data-ttu-id="f9c3e-422">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-422">Requirement</span></span>| <span data-ttu-id="f9c3e-423">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-424">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-425">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-425">1.0</span></span>|
|[<span data-ttu-id="f9c3e-426">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-427">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-428">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-429">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-430">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-430">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f9c3e-431">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-431">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f9c3e-432">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-432">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f9c3e-433">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-433">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f9c3e-434">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-434">Read mode</span></span>

<span data-ttu-id="f9c3e-435">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-435">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f9c3e-436">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f9c3e-436">Compose mode</span></span>

<span data-ttu-id="f9c3e-437">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-437">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-438">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-438">Type:</span></span>

*   <span data-ttu-id="f9c3e-439">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-439">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-440">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-440">Requirements</span></span>

|<span data-ttu-id="f9c3e-441">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-441">Requirement</span></span>| <span data-ttu-id="f9c3e-442">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-443">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-443">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-444">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-444">1.0</span></span>|
|[<span data-ttu-id="f9c3e-445">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-446">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-447">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-448">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-449">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-449">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="f9c3e-450">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-450">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="f9c3e-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f9c3e-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-455">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-455">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-456">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-456">Type:</span></span>

*   [<span data-ttu-id="f9c3e-457">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f9c3e-457">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f9c3e-458">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-458">Requirements</span></span>

|<span data-ttu-id="f9c3e-459">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-459">Requirement</span></span>| <span data-ttu-id="f9c3e-460">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-461">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-461">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-462">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-462">1.0</span></span>|
|[<span data-ttu-id="f9c3e-463">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-464">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-465">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-466">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-466">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-467">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-467">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="f9c3e-468">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-468">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="f9c3e-469">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-469">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f9c3e-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f9c3e-472">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-472">Read mode</span></span>

<span data-ttu-id="f9c3e-473">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-473">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f9c3e-474">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f9c3e-474">Compose mode</span></span>

<span data-ttu-id="f9c3e-475">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-475">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f9c3e-476">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-476">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-477">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-477">Type:</span></span>

*   <span data-ttu-id="f9c3e-478">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-478">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-479">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-479">Requirements</span></span>

|<span data-ttu-id="f9c3e-480">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-480">Requirement</span></span>| <span data-ttu-id="f9c3e-481">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-481">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-482">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-482">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-483">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-483">1.0</span></span>|
|[<span data-ttu-id="f9c3e-484">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-484">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-485">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-485">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-486">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-486">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-487">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-487">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-488">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-488">Example</span></span>

<span data-ttu-id="f9c3e-489">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-489">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="f9c3e-490">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-490">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="f9c3e-491">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f9c3e-492">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f9c3e-493">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-493">Read mode</span></span>

<span data-ttu-id="f9c3e-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="f9c3e-496">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f9c3e-496">Compose mode</span></span>

<span data-ttu-id="f9c3e-497">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f9c3e-498">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-498">Type:</span></span>

*   <span data-ttu-id="f9c3e-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-500">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-500">Requirements</span></span>

|<span data-ttu-id="f9c3e-501">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-501">Requirement</span></span>| <span data-ttu-id="f9c3e-502">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-503">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-504">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-504">1.0</span></span>|
|[<span data-ttu-id="f9c3e-505">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-506">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-507">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-508">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-508">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f9c3e-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f9c3e-510">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f9c3e-511">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f9c3e-512">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-512">Read mode</span></span>

<span data-ttu-id="f9c3e-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f9c3e-515">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f9c3e-515">Compose mode</span></span>

<span data-ttu-id="f9c3e-516">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f9c3e-517">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-517">Type:</span></span>

*   <span data-ttu-id="f9c3e-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-519">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-519">Requirements</span></span>

|<span data-ttu-id="f9c3e-520">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-520">Requirement</span></span>| <span data-ttu-id="f9c3e-521">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-522">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-523">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-523">1.0</span></span>|
|[<span data-ttu-id="f9c3e-524">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-525">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-526">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-526">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-527">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-527">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-528">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-528">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="f9c3e-529">Métodos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-529">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f9c3e-530">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f9c3e-530">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f9c3e-531">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-531">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f9c3e-532">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-532">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f9c3e-533">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-533">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f9c3e-534">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-534">Parameters:</span></span>

|<span data-ttu-id="f9c3e-535">Nome</span><span class="sxs-lookup"><span data-stu-id="f9c3e-535">Name</span></span>| <span data-ttu-id="f9c3e-536">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-536">Type</span></span>| <span data-ttu-id="f9c3e-537">Atributos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-537">Attributes</span></span>| <span data-ttu-id="f9c3e-538">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-538">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="f9c3e-539">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-539">String</span></span>||<span data-ttu-id="f9c3e-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f9c3e-542">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-542">String</span></span>||<span data-ttu-id="f9c3e-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f9c3e-545">Object</span><span class="sxs-lookup"><span data-stu-id="f9c3e-545">Object</span></span>| <span data-ttu-id="f9c3e-546">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-546">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-547">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-547">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f9c3e-548">Objeto</span><span class="sxs-lookup"><span data-stu-id="f9c3e-548">Object</span></span>| <span data-ttu-id="f9c3e-549">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-549">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-550">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-550">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f9c3e-551">function</span><span class="sxs-lookup"><span data-stu-id="f9c3e-551">function</span></span>| <span data-ttu-id="f9c3e-552">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-552">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-553">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-553">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f9c3e-554">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-554">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f9c3e-555">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-555">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f9c3e-556">Erros</span><span class="sxs-lookup"><span data-stu-id="f9c3e-556">Errors</span></span>

| <span data-ttu-id="f9c3e-557">Código de erro</span><span class="sxs-lookup"><span data-stu-id="f9c3e-557">Error code</span></span> | <span data-ttu-id="f9c3e-558">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-558">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="f9c3e-559">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-559">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="f9c3e-560">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-560">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f9c3e-561">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-561">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f9c3e-562">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-562">Requirements</span></span>

|<span data-ttu-id="f9c3e-563">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-563">Requirement</span></span>| <span data-ttu-id="f9c3e-564">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-565">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-566">1.1</span><span class="sxs-lookup"><span data-stu-id="f9c3e-566">1.1</span></span>|
|[<span data-ttu-id="f9c3e-567">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-567">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-568">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-568">ReadWriteItem</span></span>|
|[<span data-ttu-id="f9c3e-569">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-569">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-570">Escrever</span><span class="sxs-lookup"><span data-stu-id="f9c3e-570">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-571">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-571">Example</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f9c3e-572">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f9c3e-572">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f9c3e-573">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-573">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f9c3e-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f9c3e-577">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-577">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f9c3e-578">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-578">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f9c3e-579">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-579">Parameters:</span></span>

|<span data-ttu-id="f9c3e-580">Nome</span><span class="sxs-lookup"><span data-stu-id="f9c3e-580">Name</span></span>| <span data-ttu-id="f9c3e-581">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-581">Type</span></span>| <span data-ttu-id="f9c3e-582">Atributos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-582">Attributes</span></span>| <span data-ttu-id="f9c3e-583">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-583">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="f9c3e-584">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9c3e-584">String</span></span>||<span data-ttu-id="f9c3e-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f9c3e-587">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-587">String</span></span>||<span data-ttu-id="f9c3e-p136">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f9c3e-590">Object</span><span class="sxs-lookup"><span data-stu-id="f9c3e-590">Object</span></span>| <span data-ttu-id="f9c3e-591">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-591">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-592">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-592">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f9c3e-593">Objeto</span><span class="sxs-lookup"><span data-stu-id="f9c3e-593">Object</span></span>| <span data-ttu-id="f9c3e-594">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-594">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-595">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-595">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f9c3e-596">function</span><span class="sxs-lookup"><span data-stu-id="f9c3e-596">function</span></span>| <span data-ttu-id="f9c3e-597">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-597">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-598">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-598">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f9c3e-599">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-599">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f9c3e-600">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-600">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f9c3e-601">Erros</span><span class="sxs-lookup"><span data-stu-id="f9c3e-601">Errors</span></span>

| <span data-ttu-id="f9c3e-602">Código de erro</span><span class="sxs-lookup"><span data-stu-id="f9c3e-602">Error code</span></span> | <span data-ttu-id="f9c3e-603">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-603">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f9c3e-604">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-604">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f9c3e-605">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-605">Requirements</span></span>

|<span data-ttu-id="f9c3e-606">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-606">Requirement</span></span>| <span data-ttu-id="f9c3e-607">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-608">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-609">1.1</span><span class="sxs-lookup"><span data-stu-id="f9c3e-609">1.1</span></span>|
|[<span data-ttu-id="f9c3e-610">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-611">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-611">ReadWriteItem</span></span>|
|[<span data-ttu-id="f9c3e-612">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-613">Escrever</span><span class="sxs-lookup"><span data-stu-id="f9c3e-613">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-614">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-614">Example</span></span>

<span data-ttu-id="f9c3e-615">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-615">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="f9c3e-616">close()</span><span class="sxs-lookup"><span data-stu-id="f9c3e-616">close()</span></span>

<span data-ttu-id="f9c3e-617">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-617">Closes the current item that is being composed.</span></span>

<span data-ttu-id="f9c3e-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-620">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-620">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="f9c3e-621">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-621">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-622">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-622">Requirements</span></span>

|<span data-ttu-id="f9c3e-623">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-623">Requirement</span></span>| <span data-ttu-id="f9c3e-624">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-624">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-625">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-625">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-626">1.3</span><span class="sxs-lookup"><span data-stu-id="f9c3e-626">1.3</span></span>|
|[<span data-ttu-id="f9c3e-627">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-627">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-628">Restrito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-628">Restricted</span></span>|
|[<span data-ttu-id="f9c3e-629">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-629">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-630">Escrever</span><span class="sxs-lookup"><span data-stu-id="f9c3e-630">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="f9c3e-631">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-631">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="f9c3e-632">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-632">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-633">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-633">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f9c3e-634">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-634">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f9c3e-635">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-635">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="f9c3e-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f9c3e-639">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-639">Parameters:</span></span>

|<span data-ttu-id="f9c3e-640">Nome</span><span class="sxs-lookup"><span data-stu-id="f9c3e-640">Name</span></span>| <span data-ttu-id="f9c3e-641">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-641">Type</span></span>| <span data-ttu-id="f9c3e-642">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-642">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="f9c3e-643">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f9c3e-643">String &#124; Object</span></span>| |<span data-ttu-id="f9c3e-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f9c3e-646">**OU**</span><span class="sxs-lookup"><span data-stu-id="f9c3e-646">**OR**</span></span><br/><span data-ttu-id="f9c3e-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f9c3e-649">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-649">String</span></span> | <span data-ttu-id="f9c3e-650">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-650">&lt;optional&gt;</span></span> | <span data-ttu-id="f9c3e-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f9c3e-653">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-653">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f9c3e-654">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-654">&lt;optional&gt;</span></span> | <span data-ttu-id="f9c3e-655">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-655">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f9c3e-656">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-656">String</span></span> | | <span data-ttu-id="f9c3e-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f9c3e-659">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-659">String</span></span> | | <span data-ttu-id="f9c3e-660">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-660">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f9c3e-661">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-661">String</span></span> | | <span data-ttu-id="f9c3e-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f9c3e-664">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-664">String</span></span> | | <span data-ttu-id="f9c3e-p144">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f9c3e-668">function</span><span class="sxs-lookup"><span data-stu-id="f9c3e-668">function</span></span> | <span data-ttu-id="f9c3e-669">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-669">&lt;optional&gt;</span></span> | <span data-ttu-id="f9c3e-670">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-670">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f9c3e-671">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-671">Requirements</span></span>

|<span data-ttu-id="f9c3e-672">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-672">Requirement</span></span>| <span data-ttu-id="f9c3e-673">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-673">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-674">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-674">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-675">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-675">1.0</span></span>|
|[<span data-ttu-id="f9c3e-676">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-676">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-677">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-677">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-678">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-678">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-679">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-679">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f9c3e-680">Exemplos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-680">Examples</span></span>

<span data-ttu-id="f9c3e-681">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-681">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f9c3e-682">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-682">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f9c3e-683">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-683">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f9c3e-684">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-684">Reply with a body and a file attachment.</span></span>

```js
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

<span data-ttu-id="f9c3e-685">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-685">Reply with a body and an item attachment.</span></span>

```js
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

<span data-ttu-id="f9c3e-686">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-686">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="f9c3e-687">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-687">displayReplyForm(formData)</span></span>

<span data-ttu-id="f9c3e-688">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-688">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-689">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-689">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f9c3e-690">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-690">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f9c3e-691">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-691">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="f9c3e-p145">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f9c3e-695">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-695">Parameters:</span></span>

|<span data-ttu-id="f9c3e-696">Nome</span><span class="sxs-lookup"><span data-stu-id="f9c3e-696">Name</span></span>| <span data-ttu-id="f9c3e-697">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-697">Type</span></span>| <span data-ttu-id="f9c3e-698">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-698">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="f9c3e-699">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f9c3e-699">String &#124; Object</span></span>| | <span data-ttu-id="f9c3e-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f9c3e-702">**OU**</span><span class="sxs-lookup"><span data-stu-id="f9c3e-702">**OR**</span></span><br/><span data-ttu-id="f9c3e-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f9c3e-705">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-705">String</span></span> | <span data-ttu-id="f9c3e-706">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-706">&lt;optional&gt;</span></span> | <span data-ttu-id="f9c3e-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f9c3e-709">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-709">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f9c3e-710">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-710">&lt;optional&gt;</span></span> | <span data-ttu-id="f9c3e-711">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-711">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f9c3e-712">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-712">String</span></span> | | <span data-ttu-id="f9c3e-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f9c3e-715">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-715">String</span></span> | | <span data-ttu-id="f9c3e-716">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-716">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f9c3e-717">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-717">String</span></span> | | <span data-ttu-id="f9c3e-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f9c3e-720">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-720">String</span></span> | | <span data-ttu-id="f9c3e-p151">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f9c3e-724">function</span><span class="sxs-lookup"><span data-stu-id="f9c3e-724">function</span></span> | <span data-ttu-id="f9c3e-725">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-725">&lt;optional&gt;</span></span> | <span data-ttu-id="f9c3e-726">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-726">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f9c3e-727">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-727">Requirements</span></span>

|<span data-ttu-id="f9c3e-728">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-728">Requirement</span></span>| <span data-ttu-id="f9c3e-729">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-729">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-730">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-730">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-731">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-731">1.0</span></span>|
|[<span data-ttu-id="f9c3e-732">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-732">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-733">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-733">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-734">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-734">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-735">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-735">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f9c3e-736">Exemplos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-736">Examples</span></span>

<span data-ttu-id="f9c3e-737">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-737">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f9c3e-738">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-738">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f9c3e-739">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-739">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f9c3e-740">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-740">Reply with a body and a file attachment.</span></span>

```js
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

<span data-ttu-id="f9c3e-741">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-741">Reply with a body and an item attachment.</span></span>

```js
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

<span data-ttu-id="f9c3e-742">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-742">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="f9c3e-743">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f9c3e-743">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="f9c3e-744">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-744">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-745">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-745">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-746">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-746">Requirements</span></span>

|<span data-ttu-id="f9c3e-747">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-747">Requirement</span></span>| <span data-ttu-id="f9c3e-748">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-748">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-749">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-749">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-750">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-750">1.0</span></span>|
|[<span data-ttu-id="f9c3e-751">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-751">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-752">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-752">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-753">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-753">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-754">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-754">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f9c3e-755">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-755">Returns:</span></span>

<span data-ttu-id="f9c3e-756">Tipo: [Entities](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-756">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f9c3e-757">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-757">Example</span></span>

<span data-ttu-id="f9c3e-758">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-758">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="f9c3e-759">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f9c3e-759">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f9c3e-760">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-760">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-761">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-761">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f9c3e-762">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-762">Parameters:</span></span>

|<span data-ttu-id="f9c3e-763">Nome</span><span class="sxs-lookup"><span data-stu-id="f9c3e-763">Name</span></span>| <span data-ttu-id="f9c3e-764">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-764">Type</span></span>| <span data-ttu-id="f9c3e-765">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-765">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="f9c3e-766">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f9c3e-766">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="f9c3e-767">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-767">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9c3e-768">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-768">Requirements</span></span>

|<span data-ttu-id="f9c3e-769">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-769">Requirement</span></span>| <span data-ttu-id="f9c3e-770">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-770">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-771">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-771">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-772">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-772">1.0</span></span>|
|[<span data-ttu-id="f9c3e-773">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-773">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-774">Restrito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-774">Restricted</span></span>|
|[<span data-ttu-id="f9c3e-775">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-775">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-776">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-776">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f9c3e-777">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-777">Returns:</span></span>

<span data-ttu-id="f9c3e-778">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-778">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f9c3e-779">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-779">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="f9c3e-780">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-780">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f9c3e-781">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-781">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="f9c3e-782">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="f9c3e-782">Value of `entityType`</span></span> | <span data-ttu-id="f9c3e-783">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="f9c3e-783">Type of objects in returned array</span></span> | <span data-ttu-id="f9c3e-784">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="f9c3e-784">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="f9c3e-785">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-785">String</span></span> | <span data-ttu-id="f9c3e-786">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="f9c3e-786">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="f9c3e-787">Contato</span><span class="sxs-lookup"><span data-stu-id="f9c3e-787">Contact</span></span> | <span data-ttu-id="f9c3e-788">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f9c3e-788">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="f9c3e-789">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-789">String</span></span> | <span data-ttu-id="f9c3e-790">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f9c3e-790">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="f9c3e-791">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f9c3e-791">MeetingSuggestion</span></span> | <span data-ttu-id="f9c3e-792">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f9c3e-792">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="f9c3e-793">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f9c3e-793">PhoneNumber</span></span> | <span data-ttu-id="f9c3e-794">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="f9c3e-794">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="f9c3e-795">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f9c3e-795">TaskSuggestion</span></span> | <span data-ttu-id="f9c3e-796">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f9c3e-796">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="f9c3e-797">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-797">String</span></span> | <span data-ttu-id="f9c3e-798">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="f9c3e-798">**Restricted**</span></span> |

<span data-ttu-id="f9c3e-799">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f9c3e-799">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="f9c3e-800">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-800">Example</span></span>

<span data-ttu-id="f9c3e-801">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-801">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```js
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="f9c3e-802">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f9c3e-802">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f9c3e-803">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-803">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-804">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-804">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f9c3e-805">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-805">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f9c3e-806">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-806">Parameters:</span></span>

|<span data-ttu-id="f9c3e-807">Nome</span><span class="sxs-lookup"><span data-stu-id="f9c3e-807">Name</span></span>| <span data-ttu-id="f9c3e-808">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-808">Type</span></span>| <span data-ttu-id="f9c3e-809">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-809">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f9c3e-810">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-810">String</span></span>|<span data-ttu-id="f9c3e-811">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-811">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9c3e-812">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-812">Requirements</span></span>

|<span data-ttu-id="f9c3e-813">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-813">Requirement</span></span>| <span data-ttu-id="f9c3e-814">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-814">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-815">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-815">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-816">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-816">1.0</span></span>|
|[<span data-ttu-id="f9c3e-817">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-817">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-818">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-818">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-819">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-819">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-820">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-820">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f9c3e-821">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-821">Returns:</span></span>

<span data-ttu-id="f9c3e-p153">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="f9c3e-824">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f9c3e-824">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="f9c3e-825">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f9c3e-825">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f9c3e-826">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-826">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-827">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-827">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f9c3e-p154">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f9c3e-831">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-831">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f9c3e-832">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-832">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f9c3e-p155">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9c3e-836">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-836">Requirements</span></span>

|<span data-ttu-id="f9c3e-837">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-837">Requirement</span></span>| <span data-ttu-id="f9c3e-838">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-839">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-840">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-840">1.0</span></span>|
|[<span data-ttu-id="f9c3e-841">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-842">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-843">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-844">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f9c3e-845">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-845">Returns:</span></span>

<span data-ttu-id="f9c3e-p156">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="f9c3e-848">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="f9c3e-848">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f9c3e-849">Objeto</span><span class="sxs-lookup"><span data-stu-id="f9c3e-849">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f9c3e-850">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-850">Example</span></span>

<span data-ttu-id="f9c3e-851">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="f9c3e-851">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f9c3e-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="f9c3e-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f9c3e-853">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-853">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-854">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-854">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f9c3e-855">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-855">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f9c3e-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f9c3e-858">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-858">Parameters:</span></span>

|<span data-ttu-id="f9c3e-859">Nome</span><span class="sxs-lookup"><span data-stu-id="f9c3e-859">Name</span></span>| <span data-ttu-id="f9c3e-860">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-860">Type</span></span>| <span data-ttu-id="f9c3e-861">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-861">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f9c3e-862">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-862">String</span></span>|<span data-ttu-id="f9c3e-863">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-863">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9c3e-864">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-864">Requirements</span></span>

|<span data-ttu-id="f9c3e-865">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-865">Requirement</span></span>| <span data-ttu-id="f9c3e-866">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-866">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-867">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-867">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-868">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-868">1.0</span></span>|
|[<span data-ttu-id="f9c3e-869">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-869">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-870">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-870">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-871">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-871">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-872">Read</span><span class="sxs-lookup"><span data-stu-id="f9c3e-872">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f9c3e-873">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-873">Returns:</span></span>

<span data-ttu-id="f9c3e-874">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-874">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="f9c3e-875">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="f9c3e-875">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f9c3e-876">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="f9c3e-876">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f9c3e-877">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-877">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="f9c3e-878">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="f9c3e-878">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="f9c3e-879">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-879">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="f9c3e-p158">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f9c3e-882">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-882">Parameters:</span></span>

|<span data-ttu-id="f9c3e-883">Nome</span><span class="sxs-lookup"><span data-stu-id="f9c3e-883">Name</span></span>| <span data-ttu-id="f9c3e-884">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-884">Type</span></span>| <span data-ttu-id="f9c3e-885">Atributos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-885">Attributes</span></span>| <span data-ttu-id="f9c3e-886">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-886">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="f9c3e-887">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f9c3e-887">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="f9c3e-p159">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="f9c3e-891">Objeto</span><span class="sxs-lookup"><span data-stu-id="f9c3e-891">Object</span></span>| <span data-ttu-id="f9c3e-892">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-892">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-893">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-893">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f9c3e-894">Objeto</span><span class="sxs-lookup"><span data-stu-id="f9c3e-894">Object</span></span>| <span data-ttu-id="f9c3e-895">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-895">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-896">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-896">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f9c3e-897">function</span><span class="sxs-lookup"><span data-stu-id="f9c3e-897">function</span></span>||<span data-ttu-id="f9c3e-898">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-898">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f9c3e-899">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-899">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="f9c3e-900">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-900">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9c3e-901">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-901">Requirements</span></span>

|<span data-ttu-id="f9c3e-902">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-902">Requirement</span></span>| <span data-ttu-id="f9c3e-903">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-904">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-905">1.2</span><span class="sxs-lookup"><span data-stu-id="f9c3e-905">1.2</span></span>|
|[<span data-ttu-id="f9c3e-906">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-906">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-907">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-907">ReadWriteItem</span></span>|
|[<span data-ttu-id="f9c3e-908">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-908">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-909">Escrever</span><span class="sxs-lookup"><span data-stu-id="f9c3e-909">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f9c3e-910">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-910">Returns:</span></span>

<span data-ttu-id="f9c3e-911">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-911">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="f9c3e-912">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="f9c3e-912">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f9c3e-913">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-913">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f9c3e-914">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-914">Example</span></span>

```js
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f9c3e-915">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f9c3e-915">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f9c3e-916">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-916">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f9c3e-p161">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f9c3e-920">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-920">Parameters:</span></span>

|<span data-ttu-id="f9c3e-921">Nome</span><span class="sxs-lookup"><span data-stu-id="f9c3e-921">Name</span></span>| <span data-ttu-id="f9c3e-922">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-922">Type</span></span>| <span data-ttu-id="f9c3e-923">Atributos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-923">Attributes</span></span>| <span data-ttu-id="f9c3e-924">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-924">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f9c3e-925">function</span><span class="sxs-lookup"><span data-stu-id="f9c3e-925">function</span></span>||<span data-ttu-id="f9c3e-926">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-926">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f9c3e-927">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-927">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f9c3e-928">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-928">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="f9c3e-929">Objeto</span><span class="sxs-lookup"><span data-stu-id="f9c3e-929">Object</span></span>| <span data-ttu-id="f9c3e-930">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-930">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-931">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-931">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="f9c3e-932">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-932">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9c3e-933">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-933">Requirements</span></span>

|<span data-ttu-id="f9c3e-934">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-934">Requirement</span></span>| <span data-ttu-id="f9c3e-935">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-935">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-936">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-936">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-937">1.0</span><span class="sxs-lookup"><span data-stu-id="f9c3e-937">1.0</span></span>|
|[<span data-ttu-id="f9c3e-938">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-938">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-939">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-939">ReadItem</span></span>|
|[<span data-ttu-id="f9c3e-940">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-940">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-941">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9c3e-941">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-942">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-942">Example</span></span>

<span data-ttu-id="f9c3e-p164">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f9c3e-946">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f9c3e-946">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f9c3e-947">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-947">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f9c3e-p165">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f9c3e-952">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-952">Parameters:</span></span>

|<span data-ttu-id="f9c3e-953">Nome</span><span class="sxs-lookup"><span data-stu-id="f9c3e-953">Name</span></span>| <span data-ttu-id="f9c3e-954">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-954">Type</span></span>| <span data-ttu-id="f9c3e-955">Atributos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-955">Attributes</span></span>| <span data-ttu-id="f9c3e-956">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-956">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="f9c3e-957">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-957">String</span></span>||<span data-ttu-id="f9c3e-958">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-958">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="f9c3e-959">Objeto</span><span class="sxs-lookup"><span data-stu-id="f9c3e-959">Object</span></span>| <span data-ttu-id="f9c3e-960">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-960">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-961">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f9c3e-962">Objeto</span><span class="sxs-lookup"><span data-stu-id="f9c3e-962">Object</span></span>| <span data-ttu-id="f9c3e-963">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-963">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-964">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f9c3e-965">function</span><span class="sxs-lookup"><span data-stu-id="f9c3e-965">function</span></span>| <span data-ttu-id="f9c3e-966">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-966">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-967">Quando o método for concluído, a função transmitida ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f9c3e-968">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f9c3e-969">Erros</span><span class="sxs-lookup"><span data-stu-id="f9c3e-969">Errors</span></span>

| <span data-ttu-id="f9c3e-970">Código de erro</span><span class="sxs-lookup"><span data-stu-id="f9c3e-970">Error code</span></span> | <span data-ttu-id="f9c3e-971">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="f9c3e-972">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f9c3e-973">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-973">Requirements</span></span>

|<span data-ttu-id="f9c3e-974">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-974">Requirement</span></span>| <span data-ttu-id="f9c3e-975">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-976">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-976">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-977">1.1</span><span class="sxs-lookup"><span data-stu-id="f9c3e-977">1.1</span></span>|
|[<span data-ttu-id="f9c3e-978">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="f9c3e-980">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-981">Escrever</span><span class="sxs-lookup"><span data-stu-id="f9c3e-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-982">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-982">Example</span></span>

<span data-ttu-id="f9c3e-983">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-983">The following code removes an attachment with an identifier of '0'.</span></span>

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="f9c3e-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="f9c3e-985">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="f9c3e-p166">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p166">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-989">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-989">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="f9c3e-990">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="f9c3e-p168">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="f9c3e-994">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="f9c3e-995">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-995">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="f9c3e-996">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-996">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="f9c3e-997">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f9c3e-998">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-998">Parameters:</span></span>

|<span data-ttu-id="f9c3e-999">Nome</span><span class="sxs-lookup"><span data-stu-id="f9c3e-999">Name</span></span>| <span data-ttu-id="f9c3e-1000">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1000">Type</span></span>| <span data-ttu-id="f9c3e-1001">Atributos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1001">Attributes</span></span>| <span data-ttu-id="f9c3e-1002">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="f9c3e-1003">Object</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1003">Object</span></span>| <span data-ttu-id="f9c3e-1004">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-1005">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f9c3e-1006">Objeto</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1006">Object</span></span>| <span data-ttu-id="f9c3e-1007">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-1008">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1008">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="f9c3e-1009">function</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1009">function</span></span>||<span data-ttu-id="f9c3e-1010">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f9c3e-1011">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9c3e-1012">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1012">Requirements</span></span>

|<span data-ttu-id="f9c3e-1013">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1013">Requirement</span></span>| <span data-ttu-id="f9c3e-1014">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-1015">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1016">1.3</span></span>|
|[<span data-ttu-id="f9c3e-1017">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="f9c3e-1019">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-1020">Escrever</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f9c3e-1021">Exemplos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1021">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="f9c3e-p170">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="f9c3e-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="f9c3e-1025">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="f9c3e-p171">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f9c3e-1029">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1029">Parameters:</span></span>

|<span data-ttu-id="f9c3e-1030">Nome</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1030">Name</span></span>| <span data-ttu-id="f9c3e-1031">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1031">Type</span></span>| <span data-ttu-id="f9c3e-1032">Atributos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1032">Attributes</span></span>| <span data-ttu-id="f9c3e-1033">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="f9c3e-1034">String</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1034">String</span></span>||<span data-ttu-id="f9c3e-p172">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="f9c3e-1038">Objeto</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1038">Object</span></span>| <span data-ttu-id="f9c3e-1039">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-1040">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f9c3e-1041">Objeto</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1041">Object</span></span>| <span data-ttu-id="f9c3e-1042">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-1043">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="f9c3e-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="f9c3e-1045">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="f9c3e-p173">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p173">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="f9c3e-p174">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-p174">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="f9c3e-1050">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="f9c3e-1051">function</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1051">function</span></span>||<span data-ttu-id="f9c3e-1052">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f9c3e-1053">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1053">Requirements</span></span>

|<span data-ttu-id="f9c3e-1054">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1054">Requirement</span></span>| <span data-ttu-id="f9c3e-1055">Valor</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9c3e-1056">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9c3e-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1057">1.2</span></span>|
|[<span data-ttu-id="f9c3e-1058">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9c3e-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="f9c3e-1060">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9c3e-1061">Escrever</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f9c3e-1062">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f9c3e-1062">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
