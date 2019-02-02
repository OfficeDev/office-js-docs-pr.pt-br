---
title: Office.Context.Mailbox.item - requisito definir 1.1
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: ce8c10987c08609eba90a3a957b372114e62cd81
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701873"
---
# <a name="item"></a><span data-ttu-id="1fac6-102">item</span><span class="sxs-lookup"><span data-stu-id="1fac6-102">item</span></span>

### <span data-ttu-id="1fac6-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="1fac6-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="1fac6-p102">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="1fac6-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-107">Requirements</span></span>

|<span data-ttu-id="1fac6-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-108">Requirement</span></span>| <span data-ttu-id="1fac6-109">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-111">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-111">1.0</span></span>|
|[<span data-ttu-id="1fac6-112">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-113">Restrito</span><span class="sxs-lookup"><span data-stu-id="1fac6-113">Restricted</span></span>|
|[<span data-ttu-id="1fac6-114">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-115">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-115">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="1fac6-116">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-116">Example</span></span>

<span data-ttu-id="1fac6-117">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="1fac6-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
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

### <a name="members"></a><span data-ttu-id="1fac6-118">Membros</span><span class="sxs-lookup"><span data-stu-id="1fac6-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="1fac6-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1fac6-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="1fac6-p103">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-122">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="1fac6-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="1fac6-123">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="1fac6-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-124">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-124">Type:</span></span>

*   <span data-ttu-id="1fac6-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1fac6-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-126">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-126">Requirements</span></span>

|<span data-ttu-id="1fac6-127">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-127">Requirement</span></span>| <span data-ttu-id="1fac6-128">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-129">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-130">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-130">1.0</span></span>|
|[<span data-ttu-id="1fac6-131">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-132">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-134">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-135">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-135">Example</span></span>

<span data-ttu-id="1fac6-136">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="1fac6-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="1fac6-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1fac6-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="1fac6-138">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1fac6-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="1fac6-139">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="1fac6-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-140">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-140">Type:</span></span>

*   [<span data-ttu-id="1fac6-141">Destinatários</span><span class="sxs-lookup"><span data-stu-id="1fac6-141">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="1fac6-142">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-142">Requirements</span></span>

|<span data-ttu-id="1fac6-143">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-143">Requirement</span></span>| <span data-ttu-id="1fac6-144">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-145">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-146">1.1</span><span class="sxs-lookup"><span data-stu-id="1fac6-146">1.1</span></span>|
|[<span data-ttu-id="1fac6-147">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-148">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-149">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-150">Escrever</span><span class="sxs-lookup"><span data-stu-id="1fac6-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-151">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-151">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="1fac6-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="1fac6-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="1fac6-153">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="1fac6-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-154">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-154">Type:</span></span>

*   [<span data-ttu-id="1fac6-155">Corpo</span><span class="sxs-lookup"><span data-stu-id="1fac6-155">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="1fac6-156">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-156">Requirements</span></span>

|<span data-ttu-id="1fac6-157">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-157">Requirement</span></span>| <span data-ttu-id="1fac6-158">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-159">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-160">1.1</span><span class="sxs-lookup"><span data-stu-id="1fac6-160">1.1</span></span>|
|[<span data-ttu-id="1fac6-161">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-162">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-164">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-164">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="1fac6-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1fac6-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="1fac6-166">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1fac6-166">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="1fac6-167">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1fac6-167">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fac6-168">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-168">Read mode</span></span>

<span data-ttu-id="1fac6-p107">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1fac6-171">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1fac6-171">Compose mode</span></span>

<span data-ttu-id="1fac6-172">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="1fac6-172">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-173">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-173">Type:</span></span>

*   <span data-ttu-id="1fac6-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1fac6-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-175">Requirements</span></span>

|<span data-ttu-id="1fac6-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-176">Requirement</span></span>| <span data-ttu-id="1fac6-177">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-179">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-179">1.0</span></span>|
|[<span data-ttu-id="1fac6-180">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-180">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-181">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-182">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-182">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-183">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-183">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-184">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-184">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="1fac6-185">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="1fac6-185">(nullable) conversationId :String</span></span>

<span data-ttu-id="1fac6-186">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="1fac6-186">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="1fac6-p108">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="1fac6-p109">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-191">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-191">Type:</span></span>

*   <span data-ttu-id="1fac6-192">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1fac6-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-193">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-193">Requirements</span></span>

|<span data-ttu-id="1fac6-194">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-194">Requirement</span></span>| <span data-ttu-id="1fac6-195">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-196">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-197">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-197">1.0</span></span>|
|[<span data-ttu-id="1fac6-198">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-198">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-199">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-200">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-201">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-201">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="1fac6-202">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="1fac6-202">dateTimeCreated :Date</span></span>

<span data-ttu-id="1fac6-p110">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-205">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-205">Type:</span></span>

*   <span data-ttu-id="1fac6-206">Data</span><span class="sxs-lookup"><span data-stu-id="1fac6-206">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-207">Requirements</span></span>

|<span data-ttu-id="1fac6-208">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-208">Requirement</span></span>| <span data-ttu-id="1fac6-209">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-210">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-211">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-211">1.0</span></span>|
|[<span data-ttu-id="1fac6-212">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-213">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-214">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-215">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-216">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-216">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="1fac6-217">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="1fac6-217">dateTimeModified :Date</span></span>

<span data-ttu-id="1fac6-p111">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-220">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1fac6-220">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-221">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-221">Type:</span></span>

*   <span data-ttu-id="1fac6-222">Data</span><span class="sxs-lookup"><span data-stu-id="1fac6-222">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-223">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-223">Requirements</span></span>

|<span data-ttu-id="1fac6-224">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-224">Requirement</span></span>| <span data-ttu-id="1fac6-225">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-226">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-227">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-227">1.0</span></span>|
|[<span data-ttu-id="1fac6-228">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-229">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-230">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-231">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-231">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-232">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-232">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="1fac6-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="1fac6-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="1fac6-234">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="1fac6-234">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="1fac6-p112">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fac6-237">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-237">Read mode</span></span>

<span data-ttu-id="1fac6-238">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-238">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1fac6-239">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1fac6-239">Compose mode</span></span>

<span data-ttu-id="1fac6-240">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-240">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="1fac6-241">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="1fac6-241">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-242">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-242">Type:</span></span>

*   <span data-ttu-id="1fac6-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="1fac6-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-244">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-244">Requirements</span></span>

|<span data-ttu-id="1fac6-245">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-245">Requirement</span></span>| <span data-ttu-id="1fac6-246">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-247">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-248">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-248">1.0</span></span>|
|[<span data-ttu-id="1fac6-249">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-250">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-251">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-252">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-252">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-253">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-253">Example</span></span>

<span data-ttu-id="1fac6-254">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-254">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="1fac6-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1fac6-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="1fac6-p113">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="1fac6-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-260">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-260">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-261">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-261">Type:</span></span>

*   [<span data-ttu-id="1fac6-262">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1fac6-262">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1fac6-263">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-263">Requirements</span></span>

|<span data-ttu-id="1fac6-264">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-264">Requirement</span></span>| <span data-ttu-id="1fac6-265">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-266">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-267">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-267">1.0</span></span>|
|[<span data-ttu-id="1fac6-268">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-269">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-270">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-271">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-271">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="1fac6-272">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="1fac6-272">internetMessageId :String</span></span>

<span data-ttu-id="1fac6-p115">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-275">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-275">Type:</span></span>

*   <span data-ttu-id="1fac6-276">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1fac6-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-277">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-277">Requirements</span></span>

|<span data-ttu-id="1fac6-278">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-278">Requirement</span></span>| <span data-ttu-id="1fac6-279">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-280">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-281">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-281">1.0</span></span>|
|[<span data-ttu-id="1fac6-282">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-283">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-284">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-285">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-285">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-286">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-286">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="1fac6-287">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="1fac6-287">itemClass :String</span></span>

<span data-ttu-id="1fac6-p116">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="1fac6-p117">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="1fac6-292">Tipo</span><span class="sxs-lookup"><span data-stu-id="1fac6-292">Type</span></span> | <span data-ttu-id="1fac6-293">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fac6-293">Description</span></span> | <span data-ttu-id="1fac6-294">classe de item</span><span class="sxs-lookup"><span data-stu-id="1fac6-294">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="1fac6-295">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="1fac6-295">Appointment items</span></span> | <span data-ttu-id="1fac6-296">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-296">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="1fac6-297">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="1fac6-297">Message items</span></span> | <span data-ttu-id="1fac6-298">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="1fac6-298">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="1fac6-299">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="1fac6-299">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-300">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-300">Type:</span></span>

*   <span data-ttu-id="1fac6-301">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1fac6-301">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-302">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-302">Requirements</span></span>

|<span data-ttu-id="1fac6-303">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-303">Requirement</span></span>| <span data-ttu-id="1fac6-304">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-305">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-306">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-306">1.0</span></span>|
|[<span data-ttu-id="1fac6-307">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-308">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-309">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-310">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-311">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-311">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="1fac6-312">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="1fac6-312">(nullable) itemId :String</span></span>

<span data-ttu-id="1fac6-p118">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-315">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="1fac6-315">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1fac6-316">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="1fac6-316">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1fac6-317">Antes de fazer chamadas API REST usando esse valor, ele deve ser convertido usando `Office.context.mailbox.convertToRestId`, que está disponível a partir do conjunto de requisitos 1.3.</span><span class="sxs-lookup"><span data-stu-id="1fac6-317">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="1fac6-318">Para saber mais, consulte [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="1fac6-318">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-319">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-319">Type:</span></span>

*   <span data-ttu-id="1fac6-320">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1fac6-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-321">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-321">Requirements</span></span>

|<span data-ttu-id="1fac6-322">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-322">Requirement</span></span>| <span data-ttu-id="1fac6-323">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-324">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-325">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-325">1.0</span></span>|
|[<span data-ttu-id="1fac6-326">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-327">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-328">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-329">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-330">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-330">Example</span></span>

<span data-ttu-id="1fac6-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="1fac6-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="1fac6-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="1fac6-334">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="1fac6-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="1fac6-335">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="1fac6-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-336">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-336">Type:</span></span>

*   [<span data-ttu-id="1fac6-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="1fac6-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="1fac6-338">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-338">Requirements</span></span>

|<span data-ttu-id="1fac6-339">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-339">Requirement</span></span>| <span data-ttu-id="1fac6-340">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-341">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-342">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-342">1.0</span></span>|
|[<span data-ttu-id="1fac6-343">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-344">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-345">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-346">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-347">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-347">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="1fac6-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="1fac6-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="1fac6-349">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="1fac6-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fac6-350">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-350">Read mode</span></span>

<span data-ttu-id="1fac6-351">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="1fac6-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1fac6-352">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1fac6-352">Compose mode</span></span>

<span data-ttu-id="1fac6-353">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="1fac6-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-354">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-354">Type:</span></span>

*   <span data-ttu-id="1fac6-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="1fac6-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-356">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-356">Requirements</span></span>

|<span data-ttu-id="1fac6-357">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-357">Requirement</span></span>| <span data-ttu-id="1fac6-358">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-359">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-360">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-360">1.0</span></span>|
|[<span data-ttu-id="1fac6-361">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-362">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-363">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-364">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-365">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-365">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="1fac6-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="1fac6-366">normalizedSubject :String</span></span>

<span data-ttu-id="1fac6-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="1fac6-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject).</span><span class="sxs-lookup"><span data-stu-id="1fac6-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-371">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-371">Type:</span></span>

*   <span data-ttu-id="1fac6-372">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1fac6-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-373">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-373">Requirements</span></span>

|<span data-ttu-id="1fac6-374">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-374">Requirement</span></span>| <span data-ttu-id="1fac6-375">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-376">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-377">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-377">1.0</span></span>|
|[<span data-ttu-id="1fac6-378">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-379">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-380">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-381">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-382">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-382">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="1fac6-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1fac6-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="1fac6-384">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="1fac6-384">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="1fac6-385">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1fac6-385">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fac6-386">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-386">Read mode</span></span>

<span data-ttu-id="1fac6-387">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="1fac6-387">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1fac6-388">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1fac6-388">Compose mode</span></span>

<span data-ttu-id="1fac6-389">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="1fac6-389">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-390">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-390">Type:</span></span>

*   <span data-ttu-id="1fac6-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1fac6-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-392">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-392">Requirements</span></span>

|<span data-ttu-id="1fac6-393">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-393">Requirement</span></span>| <span data-ttu-id="1fac6-394">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-395">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-396">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-396">1.0</span></span>|
|[<span data-ttu-id="1fac6-397">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-398">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-399">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-400">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-400">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-401">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-401">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="1fac6-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1fac6-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="1fac6-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-405">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-405">Type:</span></span>

*   [<span data-ttu-id="1fac6-406">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1fac6-406">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1fac6-407">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-407">Requirements</span></span>

|<span data-ttu-id="1fac6-408">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-408">Requirement</span></span>| <span data-ttu-id="1fac6-409">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-410">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-411">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-411">1.0</span></span>|
|[<span data-ttu-id="1fac6-412">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-413">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-414">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-415">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-416">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-416">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="1fac6-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1fac6-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="1fac6-418">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="1fac6-418">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="1fac6-419">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1fac6-419">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fac6-420">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-420">Read mode</span></span>

<span data-ttu-id="1fac6-421">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="1fac6-421">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1fac6-422">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1fac6-422">Compose mode</span></span>

<span data-ttu-id="1fac6-423">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="1fac6-423">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-424">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-424">Type:</span></span>

*   <span data-ttu-id="1fac6-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1fac6-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-426">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-426">Requirements</span></span>

|<span data-ttu-id="1fac6-427">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-427">Requirement</span></span>| <span data-ttu-id="1fac6-428">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-429">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-430">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-430">1.0</span></span>|
|[<span data-ttu-id="1fac6-431">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-431">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-432">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-433">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-433">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-434">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-434">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-435">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-435">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="1fac6-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1fac6-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="1fac6-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="1fac6-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-441">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-441">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-442">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-442">Type:</span></span>

*   [<span data-ttu-id="1fac6-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1fac6-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1fac6-444">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-444">Requirements</span></span>

|<span data-ttu-id="1fac6-445">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-445">Requirement</span></span>| <span data-ttu-id="1fac6-446">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-447">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-448">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-448">1.0</span></span>|
|[<span data-ttu-id="1fac6-449">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-450">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-451">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-452">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-453">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-453">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="1fac6-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="1fac6-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="1fac6-455">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="1fac6-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="1fac6-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fac6-458">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-458">Read mode</span></span>

<span data-ttu-id="1fac6-459">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-459">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1fac6-460">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1fac6-460">Compose mode</span></span>

<span data-ttu-id="1fac6-461">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="1fac6-462">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="1fac6-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-463">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-463">Type:</span></span>

*   <span data-ttu-id="1fac6-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="1fac6-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-465">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-465">Requirements</span></span>

|<span data-ttu-id="1fac6-466">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-466">Requirement</span></span>| <span data-ttu-id="1fac6-467">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-468">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-469">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-469">1.0</span></span>|
|[<span data-ttu-id="1fac6-470">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-470">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-471">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-471">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-472">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-472">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-473">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-473">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-474">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-474">Example</span></span>

<span data-ttu-id="1fac6-475">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-475">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="1fac6-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1fac6-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="1fac6-477">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="1fac6-477">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="1fac6-478">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="1fac6-478">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fac6-479">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-479">Read mode</span></span>

<span data-ttu-id="1fac6-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="1fac6-482">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1fac6-482">Compose mode</span></span>

<span data-ttu-id="1fac6-483">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="1fac6-483">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1fac6-484">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-484">Type:</span></span>

*   <span data-ttu-id="1fac6-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1fac6-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-486">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-486">Requirements</span></span>

|<span data-ttu-id="1fac6-487">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-487">Requirement</span></span>| <span data-ttu-id="1fac6-488">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-488">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-489">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-490">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-490">1.0</span></span>|
|[<span data-ttu-id="1fac6-491">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-492">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-493">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-494">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-494">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="1fac6-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1fac6-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="1fac6-496">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1fac6-496">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="1fac6-497">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1fac6-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fac6-498">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-498">Read mode</span></span>

<span data-ttu-id="1fac6-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1fac6-501">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1fac6-501">Compose mode</span></span>

<span data-ttu-id="1fac6-502">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="1fac6-502">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="1fac6-503">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1fac6-503">Type:</span></span>

*   <span data-ttu-id="1fac6-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1fac6-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-505">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-505">Requirements</span></span>

|<span data-ttu-id="1fac6-506">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-506">Requirement</span></span>| <span data-ttu-id="1fac6-507">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-508">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-509">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-509">1.0</span></span>|
|[<span data-ttu-id="1fac6-510">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-511">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-512">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-513">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-514">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-514">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="1fac6-515">Métodos</span><span class="sxs-lookup"><span data-stu-id="1fac6-515">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="1fac6-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1fac6-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1fac6-517">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="1fac6-517">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1fac6-518">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="1fac6-518">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="1fac6-519">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="1fac6-519">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fac6-520">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1fac6-520">Parameters:</span></span>

|<span data-ttu-id="1fac6-521">Nome</span><span class="sxs-lookup"><span data-stu-id="1fac6-521">Name</span></span>| <span data-ttu-id="1fac6-522">Tipo</span><span class="sxs-lookup"><span data-stu-id="1fac6-522">Type</span></span>| <span data-ttu-id="1fac6-523">Atributos</span><span class="sxs-lookup"><span data-stu-id="1fac6-523">Attributes</span></span>| <span data-ttu-id="1fac6-524">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fac6-524">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="1fac6-525">String</span><span class="sxs-lookup"><span data-stu-id="1fac6-525">String</span></span>||<span data-ttu-id="1fac6-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1fac6-528">String</span><span class="sxs-lookup"><span data-stu-id="1fac6-528">String</span></span>||<span data-ttu-id="1fac6-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1fac6-531">Object</span><span class="sxs-lookup"><span data-stu-id="1fac6-531">Object</span></span>| <span data-ttu-id="1fac6-532">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-532">&lt;optional&gt;</span></span>|<span data-ttu-id="1fac6-533">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1fac6-533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1fac6-534">Objeto</span><span class="sxs-lookup"><span data-stu-id="1fac6-534">Object</span></span>| <span data-ttu-id="1fac6-535">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-535">&lt;optional&gt;</span></span>|<span data-ttu-id="1fac6-536">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1fac6-536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1fac6-537">function</span><span class="sxs-lookup"><span data-stu-id="1fac6-537">function</span></span>| <span data-ttu-id="1fac6-538">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-538">&lt;optional&gt;</span></span>|<span data-ttu-id="1fac6-539">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1fac6-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1fac6-540">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-540">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1fac6-541">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="1fac6-541">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1fac6-542">Erros</span><span class="sxs-lookup"><span data-stu-id="1fac6-542">Errors</span></span>

| <span data-ttu-id="1fac6-543">Código de erro</span><span class="sxs-lookup"><span data-stu-id="1fac6-543">Error code</span></span> | <span data-ttu-id="1fac6-544">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fac6-544">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="1fac6-545">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="1fac6-545">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="1fac6-546">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="1fac6-546">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1fac6-547">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="1fac6-547">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1fac6-548">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-548">Requirements</span></span>

|<span data-ttu-id="1fac6-549">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-549">Requirement</span></span>| <span data-ttu-id="1fac6-550">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-551">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-552">1.1</span><span class="sxs-lookup"><span data-stu-id="1fac6-552">1.1</span></span>|
|[<span data-ttu-id="1fac6-553">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-554">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-554">ReadWriteItem</span></span>|
|[<span data-ttu-id="1fac6-555">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-556">Escrever</span><span class="sxs-lookup"><span data-stu-id="1fac6-556">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-557">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-557">Example</span></span>

```JavaScript
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="1fac6-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1fac6-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1fac6-559">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="1fac6-559">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="1fac6-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="1fac6-563">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="1fac6-563">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="1fac6-564">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="1fac6-564">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fac6-565">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1fac6-565">Parameters:</span></span>

|<span data-ttu-id="1fac6-566">Nome</span><span class="sxs-lookup"><span data-stu-id="1fac6-566">Name</span></span>| <span data-ttu-id="1fac6-567">Tipo</span><span class="sxs-lookup"><span data-stu-id="1fac6-567">Type</span></span>| <span data-ttu-id="1fac6-568">Atributos</span><span class="sxs-lookup"><span data-stu-id="1fac6-568">Attributes</span></span>| <span data-ttu-id="1fac6-569">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fac6-569">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="1fac6-570">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1fac6-570">String</span></span>||<span data-ttu-id="1fac6-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1fac6-573">String</span><span class="sxs-lookup"><span data-stu-id="1fac6-573">String</span></span>||<span data-ttu-id="1fac6-p136">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1fac6-576">Object</span><span class="sxs-lookup"><span data-stu-id="1fac6-576">Object</span></span>| <span data-ttu-id="1fac6-577">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-577">&lt;optional&gt;</span></span>|<span data-ttu-id="1fac6-578">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1fac6-578">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1fac6-579">Objeto</span><span class="sxs-lookup"><span data-stu-id="1fac6-579">Object</span></span>| <span data-ttu-id="1fac6-580">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-580">&lt;optional&gt;</span></span>|<span data-ttu-id="1fac6-581">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1fac6-581">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1fac6-582">function</span><span class="sxs-lookup"><span data-stu-id="1fac6-582">function</span></span>| <span data-ttu-id="1fac6-583">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-583">&lt;optional&gt;</span></span>|<span data-ttu-id="1fac6-584">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1fac6-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1fac6-585">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-585">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1fac6-586">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="1fac6-586">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1fac6-587">Erros</span><span class="sxs-lookup"><span data-stu-id="1fac6-587">Errors</span></span>

| <span data-ttu-id="1fac6-588">Código de erro</span><span class="sxs-lookup"><span data-stu-id="1fac6-588">Error code</span></span> | <span data-ttu-id="1fac6-589">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fac6-589">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1fac6-590">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="1fac6-590">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1fac6-591">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-591">Requirements</span></span>

|<span data-ttu-id="1fac6-592">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-592">Requirement</span></span>| <span data-ttu-id="1fac6-593">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-594">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-595">1.1</span><span class="sxs-lookup"><span data-stu-id="1fac6-595">1.1</span></span>|
|[<span data-ttu-id="1fac6-596">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-596">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-597">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-597">ReadWriteItem</span></span>|
|[<span data-ttu-id="1fac6-598">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-598">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-599">Escrever</span><span class="sxs-lookup"><span data-stu-id="1fac6-599">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-600">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-600">Example</span></span>

<span data-ttu-id="1fac6-601">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-601">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="1fac6-602">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="1fac6-602">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="1fac6-603">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="1fac6-603">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-604">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1fac6-604">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1fac6-605">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="1fac6-605">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1fac6-606">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="1fac6-606">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-607">A capacidade de incluir anexos na chamada para `displayReplyAllForm` não tem suporte no conjunto de requisitos 1.1.</span><span class="sxs-lookup"><span data-stu-id="1fac6-607">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="1fac6-608">O suporte a anexos foi adicionado a `displayReplyAllForm` no conjunto de requisitos 1.2 e acima.</span><span class="sxs-lookup"><span data-stu-id="1fac6-608">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fac6-609">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1fac6-609">Parameters:</span></span>

|<span data-ttu-id="1fac6-610">Nome</span><span class="sxs-lookup"><span data-stu-id="1fac6-610">Name</span></span>| <span data-ttu-id="1fac6-611">Tipo</span><span class="sxs-lookup"><span data-stu-id="1fac6-611">Type</span></span>| <span data-ttu-id="1fac6-612">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fac6-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="1fac6-613">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1fac6-613">String &#124; Object</span></span>| |<span data-ttu-id="1fac6-p138">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1fac6-616">**OU**</span><span class="sxs-lookup"><span data-stu-id="1fac6-616">**OR**</span></span><br/><span data-ttu-id="1fac6-p139">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1fac6-619">String</span><span class="sxs-lookup"><span data-stu-id="1fac6-619">String</span></span> | <span data-ttu-id="1fac6-620">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-620">&lt;optional&gt;</span></span> | <span data-ttu-id="1fac6-p140">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="1fac6-623">function</span><span class="sxs-lookup"><span data-stu-id="1fac6-623">function</span></span> | <span data-ttu-id="1fac6-624">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-624">&lt;optional&gt;</span></span> | <span data-ttu-id="1fac6-625">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1fac6-625">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1fac6-626">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-626">Requirements</span></span>

|<span data-ttu-id="1fac6-627">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-627">Requirement</span></span>| <span data-ttu-id="1fac6-628">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-629">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-630">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-630">1.0</span></span>|
|[<span data-ttu-id="1fac6-631">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-632">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-633">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-634">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-634">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1fac6-635">Exemplos</span><span class="sxs-lookup"><span data-stu-id="1fac6-635">Examples</span></span>

<span data-ttu-id="1fac6-636">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-636">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="1fac6-637">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="1fac6-637">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="1fac6-638">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="1fac6-638">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1fac6-639">Responder com um corpo e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1fac6-639">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="1fac6-640">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="1fac6-640">displayReplyForm(formData)</span></span>

<span data-ttu-id="1fac6-641">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="1fac6-641">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-642">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1fac6-642">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1fac6-643">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="1fac6-643">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1fac6-644">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="1fac6-644">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-645">A capacidade de incluir anexos na chamada para `displayReplyForm` não tem suporte no conjunto de requisitos 1.1.</span><span class="sxs-lookup"><span data-stu-id="1fac6-645">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="1fac6-646">O suporte a anexos foi adicionado a `displayReplyForm` no conjunto de requisitos 1.2 e acima.</span><span class="sxs-lookup"><span data-stu-id="1fac6-646">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fac6-647">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1fac6-647">Parameters:</span></span>

|<span data-ttu-id="1fac6-648">Nome</span><span class="sxs-lookup"><span data-stu-id="1fac6-648">Name</span></span>| <span data-ttu-id="1fac6-649">Tipo</span><span class="sxs-lookup"><span data-stu-id="1fac6-649">Type</span></span>| <span data-ttu-id="1fac6-650">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fac6-650">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="1fac6-651">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1fac6-651">String &#124; Object</span></span>| | <span data-ttu-id="1fac6-p142">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1fac6-654">**OU**</span><span class="sxs-lookup"><span data-stu-id="1fac6-654">**OR**</span></span><br/><span data-ttu-id="1fac6-p143">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1fac6-657">String</span><span class="sxs-lookup"><span data-stu-id="1fac6-657">String</span></span> | <span data-ttu-id="1fac6-658">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-658">&lt;optional&gt;</span></span> | <span data-ttu-id="1fac6-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="1fac6-661">function</span><span class="sxs-lookup"><span data-stu-id="1fac6-661">function</span></span> | <span data-ttu-id="1fac6-662">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-662">&lt;optional&gt;</span></span> | <span data-ttu-id="1fac6-663">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1fac6-663">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1fac6-664">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-664">Requirements</span></span>

|<span data-ttu-id="1fac6-665">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-665">Requirement</span></span>| <span data-ttu-id="1fac6-666">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-666">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-667">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-667">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-668">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-668">1.0</span></span>|
|[<span data-ttu-id="1fac6-669">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-669">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-670">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-670">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-671">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-671">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-672">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-672">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1fac6-673">Exemplos</span><span class="sxs-lookup"><span data-stu-id="1fac6-673">Examples</span></span>

<span data-ttu-id="1fac6-674">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-674">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="1fac6-675">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="1fac6-675">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="1fac6-676">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="1fac6-676">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1fac6-677">Responder com um corpo e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1fac6-677">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="1fac6-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1fac6-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="1fac6-679">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="1fac6-679">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-680">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1fac6-680">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-681">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-681">Requirements</span></span>

|<span data-ttu-id="1fac6-682">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-682">Requirement</span></span>| <span data-ttu-id="1fac6-683">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-683">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-684">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-684">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-685">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-685">1.0</span></span>|
|[<span data-ttu-id="1fac6-686">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-686">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-687">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-687">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-688">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-688">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-689">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-689">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1fac6-690">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1fac6-690">Returns:</span></span>

<span data-ttu-id="1fac6-691">Tipo: [Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1fac6-691">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1fac6-692">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-692">Example</span></span>

<span data-ttu-id="1fac6-693">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1fac6-693">The following example accesses the contacts entities in the current item's body.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="1fac6-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1fac6-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1fac6-695">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="1fac6-695">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-696">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1fac6-696">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fac6-697">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1fac6-697">Parameters:</span></span>

|<span data-ttu-id="1fac6-698">Nome</span><span class="sxs-lookup"><span data-stu-id="1fac6-698">Name</span></span>| <span data-ttu-id="1fac6-699">Tipo</span><span class="sxs-lookup"><span data-stu-id="1fac6-699">Type</span></span>| <span data-ttu-id="1fac6-700">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fac6-700">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="1fac6-701">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="1fac6-701">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="1fac6-702">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="1fac6-702">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fac6-703">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-703">Requirements</span></span>

|<span data-ttu-id="1fac6-704">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-704">Requirement</span></span>| <span data-ttu-id="1fac6-705">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-705">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-706">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-706">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-707">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-707">1.0</span></span>|
|[<span data-ttu-id="1fac6-708">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-708">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-709">Restrito</span><span class="sxs-lookup"><span data-stu-id="1fac6-709">Restricted</span></span>|
|[<span data-ttu-id="1fac6-710">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-710">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-711">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-711">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1fac6-712">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1fac6-712">Returns:</span></span>

<span data-ttu-id="1fac6-713">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="1fac6-713">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="1fac6-714">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="1fac6-714">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="1fac6-715">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-715">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="1fac6-716">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="1fac6-716">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="1fac6-717">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="1fac6-717">Value of `entityType`</span></span> | <span data-ttu-id="1fac6-718">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="1fac6-718">Type of objects in returned array</span></span> | <span data-ttu-id="1fac6-719">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="1fac6-719">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="1fac6-720">String</span><span class="sxs-lookup"><span data-stu-id="1fac6-720">String</span></span> | <span data-ttu-id="1fac6-721">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="1fac6-721">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="1fac6-722">Contato</span><span class="sxs-lookup"><span data-stu-id="1fac6-722">Contact</span></span> | <span data-ttu-id="1fac6-723">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1fac6-723">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="1fac6-724">String</span><span class="sxs-lookup"><span data-stu-id="1fac6-724">String</span></span> | <span data-ttu-id="1fac6-725">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1fac6-725">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="1fac6-726">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="1fac6-726">MeetingSuggestion</span></span> | <span data-ttu-id="1fac6-727">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1fac6-727">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="1fac6-728">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="1fac6-728">PhoneNumber</span></span> | <span data-ttu-id="1fac6-729">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="1fac6-729">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="1fac6-730">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="1fac6-730">TaskSuggestion</span></span> | <span data-ttu-id="1fac6-731">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1fac6-731">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="1fac6-732">String</span><span class="sxs-lookup"><span data-stu-id="1fac6-732">String</span></span> | <span data-ttu-id="1fac6-733">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="1fac6-733">**Restricted**</span></span> |

<span data-ttu-id="1fac6-734">Tipo:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1fac6-734">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="1fac6-735">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-735">Example</span></span>

<span data-ttu-id="1fac6-736">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1fac6-736">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```JavaScript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="1fac6-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1fac6-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1fac6-738">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="1fac6-738">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-739">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1fac6-739">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1fac6-740">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="1fac6-740">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fac6-741">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1fac6-741">Parameters:</span></span>

|<span data-ttu-id="1fac6-742">Nome</span><span class="sxs-lookup"><span data-stu-id="1fac6-742">Name</span></span>| <span data-ttu-id="1fac6-743">Tipo</span><span class="sxs-lookup"><span data-stu-id="1fac6-743">Type</span></span>| <span data-ttu-id="1fac6-744">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fac6-744">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1fac6-745">String</span><span class="sxs-lookup"><span data-stu-id="1fac6-745">String</span></span>|<span data-ttu-id="1fac6-746">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="1fac6-746">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fac6-747">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-747">Requirements</span></span>

|<span data-ttu-id="1fac6-748">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-748">Requirement</span></span>| <span data-ttu-id="1fac6-749">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-749">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-750">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-750">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-751">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-751">1.0</span></span>|
|[<span data-ttu-id="1fac6-752">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-752">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-753">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-753">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-754">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-754">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-755">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-755">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1fac6-756">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1fac6-756">Returns:</span></span>

<span data-ttu-id="1fac6-p146">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="1fac6-759">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1fac6-759">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="1fac6-760">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1fac6-760">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="1fac6-761">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="1fac6-761">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-762">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1fac6-762">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1fac6-p147">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1fac6-766">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="1fac6-766">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1fac6-767">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-767">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="1fac6-p148">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fac6-770">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-770">Requirements</span></span>

|<span data-ttu-id="1fac6-771">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-771">Requirement</span></span>| <span data-ttu-id="1fac6-772">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-773">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-774">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-774">1.0</span></span>|
|[<span data-ttu-id="1fac6-775">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-775">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-776">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-776">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-777">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-777">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-778">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1fac6-779">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1fac6-779">Returns:</span></span>

<span data-ttu-id="1fac6-p149">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="1fac6-782">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="1fac6-782">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1fac6-783">Objeto</span><span class="sxs-lookup"><span data-stu-id="1fac6-783">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1fac6-784">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-784">Example</span></span>

<span data-ttu-id="1fac6-785">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="1fac6-785">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="1fac6-786">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="1fac6-786">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="1fac6-787">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="1fac6-787">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1fac6-788">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1fac6-788">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1fac6-789">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="1fac6-789">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="1fac6-p150">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fac6-792">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1fac6-792">Parameters:</span></span>

|<span data-ttu-id="1fac6-793">Nome</span><span class="sxs-lookup"><span data-stu-id="1fac6-793">Name</span></span>| <span data-ttu-id="1fac6-794">Tipo</span><span class="sxs-lookup"><span data-stu-id="1fac6-794">Type</span></span>| <span data-ttu-id="1fac6-795">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fac6-795">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1fac6-796">String</span><span class="sxs-lookup"><span data-stu-id="1fac6-796">String</span></span>|<span data-ttu-id="1fac6-797">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="1fac6-797">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fac6-798">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-798">Requirements</span></span>

|<span data-ttu-id="1fac6-799">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-799">Requirement</span></span>| <span data-ttu-id="1fac6-800">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-800">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-801">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-801">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-802">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-802">1.0</span></span>|
|[<span data-ttu-id="1fac6-803">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-803">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-804">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-804">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-805">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-805">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-806">Read</span><span class="sxs-lookup"><span data-stu-id="1fac6-806">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1fac6-807">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1fac6-807">Returns:</span></span>

<span data-ttu-id="1fac6-808">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="1fac6-808">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="1fac6-809">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="1fac6-809">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1fac6-810">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="1fac6-810">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1fac6-811">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-811">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="1fac6-812">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1fac6-812">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="1fac6-813">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="1fac6-813">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="1fac6-p151">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fac6-817">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1fac6-817">Parameters:</span></span>

|<span data-ttu-id="1fac6-818">Nome</span><span class="sxs-lookup"><span data-stu-id="1fac6-818">Name</span></span>| <span data-ttu-id="1fac6-819">Tipo</span><span class="sxs-lookup"><span data-stu-id="1fac6-819">Type</span></span>| <span data-ttu-id="1fac6-820">Atributos</span><span class="sxs-lookup"><span data-stu-id="1fac6-820">Attributes</span></span>| <span data-ttu-id="1fac6-821">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fac6-821">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1fac6-822">function</span><span class="sxs-lookup"><span data-stu-id="1fac6-822">function</span></span>||<span data-ttu-id="1fac6-823">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1fac6-823">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1fac6-824">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1fac6-824">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1fac6-825">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="1fac6-825">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="1fac6-826">Objeto</span><span class="sxs-lookup"><span data-stu-id="1fac6-826">Object</span></span>| <span data-ttu-id="1fac6-827">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-827">&lt;optional&gt;</span></span>|<span data-ttu-id="1fac6-828">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1fac6-828">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="1fac6-829">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1fac6-829">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fac6-830">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-830">Requirements</span></span>

|<span data-ttu-id="1fac6-831">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-831">Requirement</span></span>| <span data-ttu-id="1fac6-832">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-832">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-833">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-833">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-834">1.0</span><span class="sxs-lookup"><span data-stu-id="1fac6-834">1.0</span></span>|
|[<span data-ttu-id="1fac6-835">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-835">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-836">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-836">ReadItem</span></span>|
|[<span data-ttu-id="1fac6-837">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-837">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-838">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="1fac6-838">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-839">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-839">Example</span></span>

<span data-ttu-id="1fac6-p154">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="1fac6-843">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1fac6-843">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="1fac6-844">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="1fac6-844">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="1fac6-p155">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="1fac6-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fac6-849">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1fac6-849">Parameters:</span></span>

|<span data-ttu-id="1fac6-850">Nome</span><span class="sxs-lookup"><span data-stu-id="1fac6-850">Name</span></span>| <span data-ttu-id="1fac6-851">Tipo</span><span class="sxs-lookup"><span data-stu-id="1fac6-851">Type</span></span>| <span data-ttu-id="1fac6-852">Atributos</span><span class="sxs-lookup"><span data-stu-id="1fac6-852">Attributes</span></span>| <span data-ttu-id="1fac6-853">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fac6-853">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="1fac6-854">String</span><span class="sxs-lookup"><span data-stu-id="1fac6-854">String</span></span>||<span data-ttu-id="1fac6-855">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="1fac6-855">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="1fac6-856">Objeto</span><span class="sxs-lookup"><span data-stu-id="1fac6-856">Object</span></span>| <span data-ttu-id="1fac6-857">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-857">&lt;optional&gt;</span></span>|<span data-ttu-id="1fac6-858">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1fac6-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1fac6-859">Objeto</span><span class="sxs-lookup"><span data-stu-id="1fac6-859">Object</span></span>| <span data-ttu-id="1fac6-860">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-860">&lt;optional&gt;</span></span>|<span data-ttu-id="1fac6-861">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1fac6-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1fac6-862">function</span><span class="sxs-lookup"><span data-stu-id="1fac6-862">function</span></span>| <span data-ttu-id="1fac6-863">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fac6-863">&lt;optional&gt;</span></span>|<span data-ttu-id="1fac6-864">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1fac6-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1fac6-865">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="1fac6-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1fac6-866">Erros</span><span class="sxs-lookup"><span data-stu-id="1fac6-866">Errors</span></span>

| <span data-ttu-id="1fac6-867">Código de erro</span><span class="sxs-lookup"><span data-stu-id="1fac6-867">Error code</span></span> | <span data-ttu-id="1fac6-868">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fac6-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="1fac6-869">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="1fac6-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1fac6-870">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1fac6-870">Requirements</span></span>

|<span data-ttu-id="1fac6-871">Requisito</span><span class="sxs-lookup"><span data-stu-id="1fac6-871">Requirement</span></span>| <span data-ttu-id="1fac6-872">Valor</span><span class="sxs-lookup"><span data-stu-id="1fac6-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fac6-873">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1fac6-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fac6-874">1.1</span><span class="sxs-lookup"><span data-stu-id="1fac6-874">1.1</span></span>|
|[<span data-ttu-id="1fac6-875">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1fac6-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fac6-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1fac6-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="1fac6-877">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1fac6-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1fac6-878">Escrever</span><span class="sxs-lookup"><span data-stu-id="1fac6-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1fac6-879">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1fac6-879">Example</span></span>

<span data-ttu-id="1fac6-880">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="1fac6-880">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```
