---
title: Office.Context.Mailbox.item - requisito definir 1.2
description: ''
ms.date: 12/18/2018
localization_priority: Normal
ms.openlocfilehash: d58a38ce045a179a7e5cdd2e15b4e16c2ac03c91
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388595"
---
# <a name="item"></a><span data-ttu-id="300f7-102">item</span><span class="sxs-lookup"><span data-stu-id="300f7-102">item</span></span>

### <span data-ttu-id="300f7-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="300f7-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="300f7-p102">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="300f7-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-107">Requirements</span></span>

|<span data-ttu-id="300f7-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-108">Requirement</span></span>| <span data-ttu-id="300f7-109">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-111">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-111">1.0</span></span>|
|[<span data-ttu-id="300f7-112">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-113">Restrito</span><span class="sxs-lookup"><span data-stu-id="300f7-113">Restricted</span></span>|
|[<span data-ttu-id="300f7-114">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-115">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-115">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="300f7-116">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-116">Example</span></span>

<span data-ttu-id="300f7-117">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="300f7-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="300f7-118">Membros</span><span class="sxs-lookup"><span data-stu-id="300f7-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="300f7-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="300f7-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="300f7-p103">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="300f7-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="300f7-122">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="300f7-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="300f7-123">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="300f7-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-124">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-124">Type:</span></span>

*   <span data-ttu-id="300f7-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="300f7-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-126">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-126">Requirements</span></span>

|<span data-ttu-id="300f7-127">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-127">Requirement</span></span>| <span data-ttu-id="300f7-128">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-129">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-130">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-130">1.0</span></span>|
|[<span data-ttu-id="300f7-131">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-132">ReadItem</span></span>|
|[<span data-ttu-id="300f7-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-134">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-135">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-135">Example</span></span>

<span data-ttu-id="300f7-136">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="300f7-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="300f7-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="300f7-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="300f7-138">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="300f7-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="300f7-139">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="300f7-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-140">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-140">Type:</span></span>

*   [<span data-ttu-id="300f7-141">Destinatários</span><span class="sxs-lookup"><span data-stu-id="300f7-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="300f7-142">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-142">Requirements</span></span>

|<span data-ttu-id="300f7-143">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-143">Requirement</span></span>| <span data-ttu-id="300f7-144">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-145">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-146">1.1</span><span class="sxs-lookup"><span data-stu-id="300f7-146">1.1</span></span>|
|[<span data-ttu-id="300f7-147">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-148">ReadItem</span></span>|
|[<span data-ttu-id="300f7-149">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-150">Escrever</span><span class="sxs-lookup"><span data-stu-id="300f7-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-151">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-151">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="300f7-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="300f7-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="300f7-153">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="300f7-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-154">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-154">Type:</span></span>

*   [<span data-ttu-id="300f7-155">Corpo</span><span class="sxs-lookup"><span data-stu-id="300f7-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="300f7-156">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-156">Requirements</span></span>

|<span data-ttu-id="300f7-157">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-157">Requirement</span></span>| <span data-ttu-id="300f7-158">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-159">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-160">1.1</span><span class="sxs-lookup"><span data-stu-id="300f7-160">1.1</span></span>|
|[<span data-ttu-id="300f7-161">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-162">ReadItem</span></span>|
|[<span data-ttu-id="300f7-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-164">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-164">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="300f7-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="300f7-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="300f7-166">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="300f7-166">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="300f7-167">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="300f7-167">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="300f7-168">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-168">Read mode</span></span>

<span data-ttu-id="300f7-p107">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="300f7-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="300f7-171">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="300f7-171">Compose mode</span></span>

<span data-ttu-id="300f7-172">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="300f7-172">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-173">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-173">Type:</span></span>

*   <span data-ttu-id="300f7-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="300f7-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-175">Requirements</span></span>

|<span data-ttu-id="300f7-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-176">Requirement</span></span>| <span data-ttu-id="300f7-177">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-179">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-179">1.0</span></span>|
|[<span data-ttu-id="300f7-180">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-180">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-181">ReadItem</span></span>|
|[<span data-ttu-id="300f7-182">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-182">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-183">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-183">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-184">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-184">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="300f7-185">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="300f7-185">(nullable) conversationId :String</span></span>

<span data-ttu-id="300f7-186">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="300f7-186">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="300f7-p108">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="300f7-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="300f7-p109">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="300f7-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-191">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-191">Type:</span></span>

*   <span data-ttu-id="300f7-192">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="300f7-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-193">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-193">Requirements</span></span>

|<span data-ttu-id="300f7-194">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-194">Requirement</span></span>| <span data-ttu-id="300f7-195">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-196">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-197">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-197">1.0</span></span>|
|[<span data-ttu-id="300f7-198">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-198">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-199">ReadItem</span></span>|
|[<span data-ttu-id="300f7-200">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-201">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-201">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="300f7-202">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="300f7-202">dateTimeCreated :Date</span></span>

<span data-ttu-id="300f7-p110">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="300f7-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-205">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-205">Type:</span></span>

*   <span data-ttu-id="300f7-206">Data</span><span class="sxs-lookup"><span data-stu-id="300f7-206">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-207">Requirements</span></span>

|<span data-ttu-id="300f7-208">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-208">Requirement</span></span>| <span data-ttu-id="300f7-209">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-210">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-211">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-211">1.0</span></span>|
|[<span data-ttu-id="300f7-212">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-213">ReadItem</span></span>|
|[<span data-ttu-id="300f7-214">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-215">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-216">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-216">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="300f7-217">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="300f7-217">dateTimeModified :Date</span></span>

<span data-ttu-id="300f7-p111">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="300f7-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="300f7-220">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="300f7-220">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-221">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-221">Type:</span></span>

*   <span data-ttu-id="300f7-222">Data</span><span class="sxs-lookup"><span data-stu-id="300f7-222">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-223">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-223">Requirements</span></span>

|<span data-ttu-id="300f7-224">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-224">Requirement</span></span>| <span data-ttu-id="300f7-225">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-226">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-227">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-227">1.0</span></span>|
|[<span data-ttu-id="300f7-228">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-229">ReadItem</span></span>|
|[<span data-ttu-id="300f7-230">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-231">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-231">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-232">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-232">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="300f7-233">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="300f7-233">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="300f7-234">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="300f7-234">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="300f7-p112">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="300f7-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="300f7-237">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-237">Read mode</span></span>

<span data-ttu-id="300f7-238">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="300f7-238">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="300f7-239">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="300f7-239">Compose mode</span></span>

<span data-ttu-id="300f7-240">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="300f7-240">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="300f7-241">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="300f7-241">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-242">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-242">Type:</span></span>

*   <span data-ttu-id="300f7-243">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="300f7-243">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-244">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-244">Requirements</span></span>

|<span data-ttu-id="300f7-245">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-245">Requirement</span></span>| <span data-ttu-id="300f7-246">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-247">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-248">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-248">1.0</span></span>|
|[<span data-ttu-id="300f7-249">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-250">ReadItem</span></span>|
|[<span data-ttu-id="300f7-251">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-252">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-252">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-253">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-253">Example</span></span>

<span data-ttu-id="300f7-254">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="300f7-254">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="300f7-255">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="300f7-255">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="300f7-p113">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="300f7-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="300f7-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="300f7-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="300f7-260">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="300f7-260">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-261">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-261">Type:</span></span>

*   [<span data-ttu-id="300f7-262">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="300f7-262">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="300f7-263">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-263">Requirements</span></span>

|<span data-ttu-id="300f7-264">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-264">Requirement</span></span>| <span data-ttu-id="300f7-265">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-266">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-267">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-267">1.0</span></span>|
|[<span data-ttu-id="300f7-268">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-269">ReadItem</span></span>|
|[<span data-ttu-id="300f7-270">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-271">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-271">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="300f7-272">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="300f7-272">internetMessageId :String</span></span>

<span data-ttu-id="300f7-p115">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="300f7-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-275">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-275">Type:</span></span>

*   <span data-ttu-id="300f7-276">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="300f7-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-277">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-277">Requirements</span></span>

|<span data-ttu-id="300f7-278">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-278">Requirement</span></span>| <span data-ttu-id="300f7-279">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-280">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-281">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-281">1.0</span></span>|
|[<span data-ttu-id="300f7-282">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-283">ReadItem</span></span>|
|[<span data-ttu-id="300f7-284">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-285">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-285">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-286">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-286">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="300f7-287">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="300f7-287">itemClass :String</span></span>

<span data-ttu-id="300f7-p116">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="300f7-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="300f7-p117">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="300f7-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="300f7-292">Tipo</span><span class="sxs-lookup"><span data-stu-id="300f7-292">Type</span></span> | <span data-ttu-id="300f7-293">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-293">Description</span></span> | <span data-ttu-id="300f7-294">classe de item</span><span class="sxs-lookup"><span data-stu-id="300f7-294">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="300f7-295">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="300f7-295">Appointment items</span></span> | <span data-ttu-id="300f7-296">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="300f7-296">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="300f7-297">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="300f7-297">Message items</span></span> | <span data-ttu-id="300f7-298">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="300f7-298">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="300f7-299">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="300f7-299">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-300">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-300">Type:</span></span>

*   <span data-ttu-id="300f7-301">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="300f7-301">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-302">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-302">Requirements</span></span>

|<span data-ttu-id="300f7-303">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-303">Requirement</span></span>| <span data-ttu-id="300f7-304">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-305">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-306">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-306">1.0</span></span>|
|[<span data-ttu-id="300f7-307">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-308">ReadItem</span></span>|
|[<span data-ttu-id="300f7-309">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-310">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-311">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-311">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="300f7-312">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="300f7-312">(nullable) itemId :String</span></span>

<span data-ttu-id="300f7-p118">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="300f7-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="300f7-315">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="300f7-315">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="300f7-316">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="300f7-316">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="300f7-317">Antes de fazer chamadas API REST usando esse valor, ele deve ser convertido usando `Office.context.mailbox.convertToRestId`, que está disponível a partir do conjunto de requisitos 1.3.</span><span class="sxs-lookup"><span data-stu-id="300f7-317">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="300f7-318">Para saber mais, consulte [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="300f7-318">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-319">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-319">Type:</span></span>

*   <span data-ttu-id="300f7-320">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="300f7-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-321">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-321">Requirements</span></span>

|<span data-ttu-id="300f7-322">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-322">Requirement</span></span>| <span data-ttu-id="300f7-323">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-324">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-325">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-325">1.0</span></span>|
|[<span data-ttu-id="300f7-326">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-327">ReadItem</span></span>|
|[<span data-ttu-id="300f7-328">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-329">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-330">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-330">Example</span></span>

<span data-ttu-id="300f7-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="300f7-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="300f7-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="300f7-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="300f7-334">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="300f7-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="300f7-335">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="300f7-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-336">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-336">Type:</span></span>

*   [<span data-ttu-id="300f7-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="300f7-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="300f7-338">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-338">Requirements</span></span>

|<span data-ttu-id="300f7-339">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-339">Requirement</span></span>| <span data-ttu-id="300f7-340">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-341">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-342">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-342">1.0</span></span>|
|[<span data-ttu-id="300f7-343">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-344">ReadItem</span></span>|
|[<span data-ttu-id="300f7-345">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-346">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-347">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-347">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="300f7-348">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="300f7-348">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="300f7-349">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="300f7-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="300f7-350">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-350">Read mode</span></span>

<span data-ttu-id="300f7-351">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="300f7-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="300f7-352">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="300f7-352">Compose mode</span></span>

<span data-ttu-id="300f7-353">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="300f7-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-354">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-354">Type:</span></span>

*   <span data-ttu-id="300f7-355">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="300f7-355">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-356">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-356">Requirements</span></span>

|<span data-ttu-id="300f7-357">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-357">Requirement</span></span>| <span data-ttu-id="300f7-358">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-359">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-360">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-360">1.0</span></span>|
|[<span data-ttu-id="300f7-361">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-362">ReadItem</span></span>|
|[<span data-ttu-id="300f7-363">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-364">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-365">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-365">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="300f7-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="300f7-366">normalizedSubject :String</span></span>

<span data-ttu-id="300f7-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="300f7-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="300f7-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject).</span><span class="sxs-lookup"><span data-stu-id="300f7-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-371">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-371">Type:</span></span>

*   <span data-ttu-id="300f7-372">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="300f7-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-373">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-373">Requirements</span></span>

|<span data-ttu-id="300f7-374">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-374">Requirement</span></span>| <span data-ttu-id="300f7-375">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-376">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-377">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-377">1.0</span></span>|
|[<span data-ttu-id="300f7-378">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-379">ReadItem</span></span>|
|[<span data-ttu-id="300f7-380">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-381">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-382">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-382">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="300f7-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="300f7-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="300f7-384">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="300f7-384">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="300f7-385">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="300f7-385">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="300f7-386">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-386">Read mode</span></span>

<span data-ttu-id="300f7-387">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="300f7-387">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="300f7-388">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="300f7-388">Compose mode</span></span>

<span data-ttu-id="300f7-389">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="300f7-389">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-390">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-390">Type:</span></span>

*   <span data-ttu-id="300f7-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="300f7-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-392">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-392">Requirements</span></span>

|<span data-ttu-id="300f7-393">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-393">Requirement</span></span>| <span data-ttu-id="300f7-394">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-395">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-396">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-396">1.0</span></span>|
|[<span data-ttu-id="300f7-397">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-398">ReadItem</span></span>|
|[<span data-ttu-id="300f7-399">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-400">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-400">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-401">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-401">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="300f7-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="300f7-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="300f7-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="300f7-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-405">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-405">Type:</span></span>

*   [<span data-ttu-id="300f7-406">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="300f7-406">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="300f7-407">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-407">Requirements</span></span>

|<span data-ttu-id="300f7-408">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-408">Requirement</span></span>| <span data-ttu-id="300f7-409">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-410">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-411">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-411">1.0</span></span>|
|[<span data-ttu-id="300f7-412">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-413">ReadItem</span></span>|
|[<span data-ttu-id="300f7-414">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-415">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-416">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-416">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="300f7-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="300f7-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="300f7-418">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="300f7-418">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="300f7-419">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="300f7-419">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="300f7-420">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-420">Read mode</span></span>

<span data-ttu-id="300f7-421">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="300f7-421">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="300f7-422">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="300f7-422">Compose mode</span></span>

<span data-ttu-id="300f7-423">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="300f7-423">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-424">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-424">Type:</span></span>

*   <span data-ttu-id="300f7-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="300f7-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-426">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-426">Requirements</span></span>

|<span data-ttu-id="300f7-427">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-427">Requirement</span></span>| <span data-ttu-id="300f7-428">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-429">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-430">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-430">1.0</span></span>|
|[<span data-ttu-id="300f7-431">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-431">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-432">ReadItem</span></span>|
|[<span data-ttu-id="300f7-433">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-433">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-434">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-434">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-435">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-435">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="300f7-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="300f7-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="300f7-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="300f7-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="300f7-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="300f7-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="300f7-441">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="300f7-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-442">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-442">Type:</span></span>

*   [<span data-ttu-id="300f7-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="300f7-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="300f7-444">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-444">Requirements</span></span>

|<span data-ttu-id="300f7-445">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-445">Requirement</span></span>| <span data-ttu-id="300f7-446">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-447">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-448">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-448">1.0</span></span>|
|[<span data-ttu-id="300f7-449">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-450">ReadItem</span></span>|
|[<span data-ttu-id="300f7-451">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-452">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-453">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-453">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="300f7-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="300f7-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="300f7-455">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="300f7-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="300f7-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="300f7-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="300f7-458">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-458">Read mode</span></span>

<span data-ttu-id="300f7-459">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="300f7-459">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="300f7-460">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="300f7-460">Compose mode</span></span>

<span data-ttu-id="300f7-461">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="300f7-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="300f7-462">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="300f7-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-463">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-463">Type:</span></span>

*   <span data-ttu-id="300f7-464">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="300f7-464">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-465">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-465">Requirements</span></span>

|<span data-ttu-id="300f7-466">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-466">Requirement</span></span>| <span data-ttu-id="300f7-467">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-468">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-469">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-469">1.0</span></span>|
|[<span data-ttu-id="300f7-470">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-470">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-471">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-471">ReadItem</span></span>|
|[<span data-ttu-id="300f7-472">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-472">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-473">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-473">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-474">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-474">Example</span></span>

<span data-ttu-id="300f7-475">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="300f7-475">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="300f7-476">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="300f7-476">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="300f7-477">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="300f7-477">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="300f7-478">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="300f7-478">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="300f7-479">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-479">Read mode</span></span>

<span data-ttu-id="300f7-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="300f7-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="300f7-482">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="300f7-482">Compose mode</span></span>

<span data-ttu-id="300f7-483">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="300f7-483">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="300f7-484">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-484">Type:</span></span>

*   <span data-ttu-id="300f7-485">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="300f7-485">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-486">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-486">Requirements</span></span>

|<span data-ttu-id="300f7-487">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-487">Requirement</span></span>| <span data-ttu-id="300f7-488">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-488">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-489">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-490">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-490">1.0</span></span>|
|[<span data-ttu-id="300f7-491">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-492">ReadItem</span></span>|
|[<span data-ttu-id="300f7-493">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-494">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-494">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="300f7-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="300f7-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="300f7-496">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="300f7-496">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="300f7-497">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="300f7-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="300f7-498">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-498">Read mode</span></span>

<span data-ttu-id="300f7-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="300f7-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="300f7-501">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="300f7-501">Compose mode</span></span>

<span data-ttu-id="300f7-502">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="300f7-502">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="300f7-503">Tipo:</span><span class="sxs-lookup"><span data-stu-id="300f7-503">Type:</span></span>

*   <span data-ttu-id="300f7-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="300f7-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-505">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-505">Requirements</span></span>

|<span data-ttu-id="300f7-506">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-506">Requirement</span></span>| <span data-ttu-id="300f7-507">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-508">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-509">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-509">1.0</span></span>|
|[<span data-ttu-id="300f7-510">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-511">ReadItem</span></span>|
|[<span data-ttu-id="300f7-512">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-513">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-514">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-514">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="300f7-515">Métodos</span><span class="sxs-lookup"><span data-stu-id="300f7-515">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="300f7-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="300f7-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="300f7-517">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="300f7-517">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="300f7-518">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="300f7-518">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="300f7-519">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="300f7-519">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="300f7-520">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="300f7-520">Parameters:</span></span>

|<span data-ttu-id="300f7-521">Nome</span><span class="sxs-lookup"><span data-stu-id="300f7-521">Name</span></span>| <span data-ttu-id="300f7-522">Tipo</span><span class="sxs-lookup"><span data-stu-id="300f7-522">Type</span></span>| <span data-ttu-id="300f7-523">Atributos</span><span class="sxs-lookup"><span data-stu-id="300f7-523">Attributes</span></span>| <span data-ttu-id="300f7-524">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-524">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="300f7-525">String</span><span class="sxs-lookup"><span data-stu-id="300f7-525">String</span></span>||<span data-ttu-id="300f7-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="300f7-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="300f7-528">String</span><span class="sxs-lookup"><span data-stu-id="300f7-528">String</span></span>||<span data-ttu-id="300f7-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="300f7-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="300f7-531">Objeto</span><span class="sxs-lookup"><span data-stu-id="300f7-531">Object</span></span>| <span data-ttu-id="300f7-532">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-532">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-533">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="300f7-533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="300f7-534">Objeto</span><span class="sxs-lookup"><span data-stu-id="300f7-534">Object</span></span>| <span data-ttu-id="300f7-535">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-535">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-536">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="300f7-536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="300f7-537">function</span><span class="sxs-lookup"><span data-stu-id="300f7-537">function</span></span>| <span data-ttu-id="300f7-538">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-538">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-539">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="300f7-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="300f7-540">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="300f7-540">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="300f7-541">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="300f7-541">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="300f7-542">Erros</span><span class="sxs-lookup"><span data-stu-id="300f7-542">Errors</span></span>

| <span data-ttu-id="300f7-543">Código de erro</span><span class="sxs-lookup"><span data-stu-id="300f7-543">Error code</span></span> | <span data-ttu-id="300f7-544">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-544">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="300f7-545">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="300f7-545">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="300f7-546">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="300f7-546">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="300f7-547">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="300f7-547">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="300f7-548">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-548">Requirements</span></span>

|<span data-ttu-id="300f7-549">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-549">Requirement</span></span>| <span data-ttu-id="300f7-550">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-551">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-552">1.1</span><span class="sxs-lookup"><span data-stu-id="300f7-552">1.1</span></span>|
|[<span data-ttu-id="300f7-553">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-554">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="300f7-554">ReadWriteItem</span></span>|
|[<span data-ttu-id="300f7-555">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-556">Escrever</span><span class="sxs-lookup"><span data-stu-id="300f7-556">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-557">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-557">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="300f7-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="300f7-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="300f7-559">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="300f7-559">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="300f7-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="300f7-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="300f7-563">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="300f7-563">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="300f7-564">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="300f7-564">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="300f7-565">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="300f7-565">Parameters:</span></span>

|<span data-ttu-id="300f7-566">Nome</span><span class="sxs-lookup"><span data-stu-id="300f7-566">Name</span></span>| <span data-ttu-id="300f7-567">Tipo</span><span class="sxs-lookup"><span data-stu-id="300f7-567">Type</span></span>| <span data-ttu-id="300f7-568">Atributos</span><span class="sxs-lookup"><span data-stu-id="300f7-568">Attributes</span></span>| <span data-ttu-id="300f7-569">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-569">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="300f7-570">String</span><span class="sxs-lookup"><span data-stu-id="300f7-570">String</span></span>||<span data-ttu-id="300f7-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="300f7-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="300f7-573">String</span><span class="sxs-lookup"><span data-stu-id="300f7-573">String</span></span>||<span data-ttu-id="300f7-p136">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="300f7-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="300f7-576">Objeto</span><span class="sxs-lookup"><span data-stu-id="300f7-576">Object</span></span>| <span data-ttu-id="300f7-577">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-577">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-578">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="300f7-578">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="300f7-579">Objeto</span><span class="sxs-lookup"><span data-stu-id="300f7-579">Object</span></span>| <span data-ttu-id="300f7-580">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-580">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-581">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="300f7-581">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="300f7-582">function</span><span class="sxs-lookup"><span data-stu-id="300f7-582">function</span></span>| <span data-ttu-id="300f7-583">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-583">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-584">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="300f7-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="300f7-585">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="300f7-585">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="300f7-586">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="300f7-586">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="300f7-587">Erros</span><span class="sxs-lookup"><span data-stu-id="300f7-587">Errors</span></span>

| <span data-ttu-id="300f7-588">Código de erro</span><span class="sxs-lookup"><span data-stu-id="300f7-588">Error code</span></span> | <span data-ttu-id="300f7-589">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-589">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="300f7-590">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="300f7-590">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="300f7-591">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-591">Requirements</span></span>

|<span data-ttu-id="300f7-592">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-592">Requirement</span></span>| <span data-ttu-id="300f7-593">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-594">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-595">1.1</span><span class="sxs-lookup"><span data-stu-id="300f7-595">1.1</span></span>|
|[<span data-ttu-id="300f7-596">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-596">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-597">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="300f7-597">ReadWriteItem</span></span>|
|[<span data-ttu-id="300f7-598">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-598">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-599">Escrever</span><span class="sxs-lookup"><span data-stu-id="300f7-599">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-600">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-600">Example</span></span>

<span data-ttu-id="300f7-601">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="300f7-601">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="300f7-602">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="300f7-602">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="300f7-603">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="300f7-603">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="300f7-604">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="300f7-604">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="300f7-605">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="300f7-605">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="300f7-606">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="300f7-606">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="300f7-p137">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="300f7-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="300f7-610">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="300f7-610">Parameters:</span></span>

|<span data-ttu-id="300f7-611">Nome</span><span class="sxs-lookup"><span data-stu-id="300f7-611">Name</span></span>| <span data-ttu-id="300f7-612">Tipo</span><span class="sxs-lookup"><span data-stu-id="300f7-612">Type</span></span>| <span data-ttu-id="300f7-613">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-613">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="300f7-614">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="300f7-614">String &#124; Object</span></span>| |<span data-ttu-id="300f7-p138">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="300f7-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="300f7-617">**OU**</span><span class="sxs-lookup"><span data-stu-id="300f7-617">**OR**</span></span><br/><span data-ttu-id="300f7-p139">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="300f7-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="300f7-620">String</span><span class="sxs-lookup"><span data-stu-id="300f7-620">String</span></span> | <span data-ttu-id="300f7-621">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-621">&lt;optional&gt;</span></span> | <span data-ttu-id="300f7-p140">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="300f7-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="300f7-624">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-624">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="300f7-625">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-625">&lt;optional&gt;</span></span> | <span data-ttu-id="300f7-626">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="300f7-626">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="300f7-627">String</span><span class="sxs-lookup"><span data-stu-id="300f7-627">String</span></span> | | <span data-ttu-id="300f7-p141">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="300f7-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="300f7-630">String</span><span class="sxs-lookup"><span data-stu-id="300f7-630">String</span></span> | | <span data-ttu-id="300f7-631">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="300f7-631">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="300f7-632">String</span><span class="sxs-lookup"><span data-stu-id="300f7-632">String</span></span> | | <span data-ttu-id="300f7-p142">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="300f7-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="300f7-635">String</span><span class="sxs-lookup"><span data-stu-id="300f7-635">String</span></span> | | <span data-ttu-id="300f7-p143">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="300f7-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="300f7-639">function</span><span class="sxs-lookup"><span data-stu-id="300f7-639">function</span></span> | <span data-ttu-id="300f7-640">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-640">&lt;optional&gt;</span></span> | <span data-ttu-id="300f7-641">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="300f7-641">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="300f7-642">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-642">Requirements</span></span>

|<span data-ttu-id="300f7-643">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-643">Requirement</span></span>| <span data-ttu-id="300f7-644">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-644">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-645">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-645">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-646">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-646">1.0</span></span>|
|[<span data-ttu-id="300f7-647">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-647">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-648">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-648">ReadItem</span></span>|
|[<span data-ttu-id="300f7-649">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-649">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-650">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-650">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="300f7-651">Exemplos</span><span class="sxs-lookup"><span data-stu-id="300f7-651">Examples</span></span>

<span data-ttu-id="300f7-652">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="300f7-652">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="300f7-653">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="300f7-653">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="300f7-654">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="300f7-654">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="300f7-655">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="300f7-655">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="300f7-656">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="300f7-656">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="300f7-657">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="300f7-657">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="300f7-658">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="300f7-658">displayReplyForm(formData)</span></span>

<span data-ttu-id="300f7-659">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="300f7-659">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="300f7-660">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="300f7-660">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="300f7-661">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="300f7-661">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="300f7-662">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="300f7-662">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="300f7-p144">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="300f7-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="300f7-666">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="300f7-666">Parameters:</span></span>

|<span data-ttu-id="300f7-667">Nome</span><span class="sxs-lookup"><span data-stu-id="300f7-667">Name</span></span>| <span data-ttu-id="300f7-668">Tipo</span><span class="sxs-lookup"><span data-stu-id="300f7-668">Type</span></span>| <span data-ttu-id="300f7-669">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-669">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="300f7-670">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="300f7-670">String &#124; Object</span></span>| | <span data-ttu-id="300f7-p145">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="300f7-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="300f7-673">**OU**</span><span class="sxs-lookup"><span data-stu-id="300f7-673">**OR**</span></span><br/><span data-ttu-id="300f7-p146">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="300f7-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="300f7-676">String</span><span class="sxs-lookup"><span data-stu-id="300f7-676">String</span></span> | <span data-ttu-id="300f7-677">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-677">&lt;optional&gt;</span></span> | <span data-ttu-id="300f7-p147">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="300f7-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="300f7-680">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-680">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="300f7-681">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-681">&lt;optional&gt;</span></span> | <span data-ttu-id="300f7-682">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="300f7-682">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="300f7-683">String</span><span class="sxs-lookup"><span data-stu-id="300f7-683">String</span></span> | | <span data-ttu-id="300f7-p148">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="300f7-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="300f7-686">String</span><span class="sxs-lookup"><span data-stu-id="300f7-686">String</span></span> | | <span data-ttu-id="300f7-687">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="300f7-687">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="300f7-688">String</span><span class="sxs-lookup"><span data-stu-id="300f7-688">String</span></span> | | <span data-ttu-id="300f7-p149">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="300f7-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="300f7-691">String</span><span class="sxs-lookup"><span data-stu-id="300f7-691">String</span></span> | | <span data-ttu-id="300f7-p150">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="300f7-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="300f7-695">function</span><span class="sxs-lookup"><span data-stu-id="300f7-695">function</span></span> | <span data-ttu-id="300f7-696">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-696">&lt;optional&gt;</span></span> | <span data-ttu-id="300f7-697">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="300f7-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="300f7-698">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-698">Requirements</span></span>

|<span data-ttu-id="300f7-699">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-699">Requirement</span></span>| <span data-ttu-id="300f7-700">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-700">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-701">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-701">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-702">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-702">1.0</span></span>|
|[<span data-ttu-id="300f7-703">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-703">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-704">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-704">ReadItem</span></span>|
|[<span data-ttu-id="300f7-705">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-705">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-706">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-706">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="300f7-707">Exemplos</span><span class="sxs-lookup"><span data-stu-id="300f7-707">Examples</span></span>

<span data-ttu-id="300f7-708">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="300f7-708">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="300f7-709">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="300f7-709">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="300f7-710">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="300f7-710">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="300f7-711">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="300f7-711">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="300f7-712">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="300f7-712">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="300f7-713">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="300f7-713">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="300f7-714">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="300f7-714">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="300f7-715">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="300f7-715">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="300f7-716">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="300f7-716">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-717">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-717">Requirements</span></span>

|<span data-ttu-id="300f7-718">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-718">Requirement</span></span>| <span data-ttu-id="300f7-719">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-719">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-720">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-720">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-721">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-721">1.0</span></span>|
|[<span data-ttu-id="300f7-722">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-722">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-723">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-723">ReadItem</span></span>|
|[<span data-ttu-id="300f7-724">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-724">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-725">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-725">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="300f7-726">Retorna:</span><span class="sxs-lookup"><span data-stu-id="300f7-726">Returns:</span></span>

<span data-ttu-id="300f7-727">Tipo: [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="300f7-727">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="300f7-728">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-728">Example</span></span>

<span data-ttu-id="300f7-729">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="300f7-729">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="300f7-730">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="300f7-730">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="300f7-731">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="300f7-731">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="300f7-732">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="300f7-732">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="300f7-733">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="300f7-733">Parameters:</span></span>

|<span data-ttu-id="300f7-734">Nome</span><span class="sxs-lookup"><span data-stu-id="300f7-734">Name</span></span>| <span data-ttu-id="300f7-735">Tipo</span><span class="sxs-lookup"><span data-stu-id="300f7-735">Type</span></span>| <span data-ttu-id="300f7-736">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-736">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="300f7-737">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="300f7-737">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="300f7-738">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="300f7-738">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="300f7-739">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-739">Requirements</span></span>

|<span data-ttu-id="300f7-740">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-740">Requirement</span></span>| <span data-ttu-id="300f7-741">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-742">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-743">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-743">1.0</span></span>|
|[<span data-ttu-id="300f7-744">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-744">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-745">Restrito</span><span class="sxs-lookup"><span data-stu-id="300f7-745">Restricted</span></span>|
|[<span data-ttu-id="300f7-746">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-746">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-747">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-747">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="300f7-748">Retorna:</span><span class="sxs-lookup"><span data-stu-id="300f7-748">Returns:</span></span>

<span data-ttu-id="300f7-749">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="300f7-749">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="300f7-750">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="300f7-750">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="300f7-751">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="300f7-751">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="300f7-752">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="300f7-752">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="300f7-753">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="300f7-753">Value of `entityType`</span></span> | <span data-ttu-id="300f7-754">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="300f7-754">Type of objects in returned array</span></span> | <span data-ttu-id="300f7-755">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="300f7-755">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="300f7-756">String</span><span class="sxs-lookup"><span data-stu-id="300f7-756">String</span></span> | <span data-ttu-id="300f7-757">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="300f7-757">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="300f7-758">Contato</span><span class="sxs-lookup"><span data-stu-id="300f7-758">Contact</span></span> | <span data-ttu-id="300f7-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="300f7-759">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="300f7-760">String</span><span class="sxs-lookup"><span data-stu-id="300f7-760">String</span></span> | <span data-ttu-id="300f7-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="300f7-761">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="300f7-762">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="300f7-762">MeetingSuggestion</span></span> | <span data-ttu-id="300f7-763">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="300f7-763">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="300f7-764">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="300f7-764">PhoneNumber</span></span> | <span data-ttu-id="300f7-765">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="300f7-765">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="300f7-766">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="300f7-766">TaskSuggestion</span></span> | <span data-ttu-id="300f7-767">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="300f7-767">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="300f7-768">String</span><span class="sxs-lookup"><span data-stu-id="300f7-768">String</span></span> | <span data-ttu-id="300f7-769">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="300f7-769">**Restricted**</span></span> |

<span data-ttu-id="300f7-770">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="300f7-770">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="300f7-771">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-771">Example</span></span>

<span data-ttu-id="300f7-772">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="300f7-772">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="300f7-773">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="300f7-773">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="300f7-774">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="300f7-774">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="300f7-775">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="300f7-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="300f7-776">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="300f7-776">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="300f7-777">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="300f7-777">Parameters:</span></span>

|<span data-ttu-id="300f7-778">Nome</span><span class="sxs-lookup"><span data-stu-id="300f7-778">Name</span></span>| <span data-ttu-id="300f7-779">Tipo</span><span class="sxs-lookup"><span data-stu-id="300f7-779">Type</span></span>| <span data-ttu-id="300f7-780">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-780">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="300f7-781">String</span><span class="sxs-lookup"><span data-stu-id="300f7-781">String</span></span>|<span data-ttu-id="300f7-782">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="300f7-782">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="300f7-783">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-783">Requirements</span></span>

|<span data-ttu-id="300f7-784">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-784">Requirement</span></span>| <span data-ttu-id="300f7-785">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-786">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-786">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-787">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-787">1.0</span></span>|
|[<span data-ttu-id="300f7-788">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-789">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-789">ReadItem</span></span>|
|[<span data-ttu-id="300f7-790">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-791">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-791">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="300f7-792">Retorna:</span><span class="sxs-lookup"><span data-stu-id="300f7-792">Returns:</span></span>

<span data-ttu-id="300f7-p152">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="300f7-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="300f7-795">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="300f7-795">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="300f7-796">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="300f7-796">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="300f7-797">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="300f7-797">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="300f7-798">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="300f7-798">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="300f7-p153">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="300f7-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="300f7-802">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="300f7-802">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="300f7-803">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="300f7-803">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="300f7-p154">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="300f7-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="300f7-806">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-806">Requirements</span></span>

|<span data-ttu-id="300f7-807">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-807">Requirement</span></span>| <span data-ttu-id="300f7-808">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-808">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-809">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-809">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-810">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-810">1.0</span></span>|
|[<span data-ttu-id="300f7-811">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-811">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-812">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-812">ReadItem</span></span>|
|[<span data-ttu-id="300f7-813">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-813">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-814">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-814">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="300f7-815">Retorna:</span><span class="sxs-lookup"><span data-stu-id="300f7-815">Returns:</span></span>

<span data-ttu-id="300f7-p155">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="300f7-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="300f7-818">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="300f7-818">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="300f7-819">Objeto</span><span class="sxs-lookup"><span data-stu-id="300f7-819">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="300f7-820">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-820">Example</span></span>

<span data-ttu-id="300f7-821">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="300f7-821">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="300f7-822">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="300f7-822">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="300f7-823">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="300f7-823">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="300f7-824">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="300f7-824">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="300f7-825">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="300f7-825">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="300f7-p156">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="300f7-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="300f7-828">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="300f7-828">Parameters:</span></span>

|<span data-ttu-id="300f7-829">Nome</span><span class="sxs-lookup"><span data-stu-id="300f7-829">Name</span></span>| <span data-ttu-id="300f7-830">Tipo</span><span class="sxs-lookup"><span data-stu-id="300f7-830">Type</span></span>| <span data-ttu-id="300f7-831">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-831">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="300f7-832">String</span><span class="sxs-lookup"><span data-stu-id="300f7-832">String</span></span>|<span data-ttu-id="300f7-833">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="300f7-833">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="300f7-834">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-834">Requirements</span></span>

|<span data-ttu-id="300f7-835">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-835">Requirement</span></span>| <span data-ttu-id="300f7-836">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-836">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-837">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-837">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-838">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-838">1.0</span></span>|
|[<span data-ttu-id="300f7-839">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-839">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-840">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-840">ReadItem</span></span>|
|[<span data-ttu-id="300f7-841">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-841">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-842">Read</span><span class="sxs-lookup"><span data-stu-id="300f7-842">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="300f7-843">Retorna:</span><span class="sxs-lookup"><span data-stu-id="300f7-843">Returns:</span></span>

<span data-ttu-id="300f7-844">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="300f7-844">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="300f7-845">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="300f7-845">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="300f7-846">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="300f7-846">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="300f7-847">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-847">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="300f7-848">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="300f7-848">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="300f7-849">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="300f7-849">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="300f7-p157">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="300f7-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="300f7-852">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="300f7-852">Parameters:</span></span>

|<span data-ttu-id="300f7-853">Nome</span><span class="sxs-lookup"><span data-stu-id="300f7-853">Name</span></span>| <span data-ttu-id="300f7-854">Tipo</span><span class="sxs-lookup"><span data-stu-id="300f7-854">Type</span></span>| <span data-ttu-id="300f7-855">Atributos</span><span class="sxs-lookup"><span data-stu-id="300f7-855">Attributes</span></span>| <span data-ttu-id="300f7-856">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-856">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="300f7-857">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="300f7-857">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="300f7-p158">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="300f7-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="300f7-861">Objeto</span><span class="sxs-lookup"><span data-stu-id="300f7-861">Object</span></span>| <span data-ttu-id="300f7-862">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-862">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-863">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="300f7-863">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="300f7-864">Objeto</span><span class="sxs-lookup"><span data-stu-id="300f7-864">Object</span></span>| <span data-ttu-id="300f7-865">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-865">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-866">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="300f7-866">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="300f7-867">function</span><span class="sxs-lookup"><span data-stu-id="300f7-867">function</span></span>||<span data-ttu-id="300f7-868">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="300f7-868">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="300f7-869">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="300f7-869">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="300f7-870">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="300f7-870">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="300f7-871">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-871">Requirements</span></span>

|<span data-ttu-id="300f7-872">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-872">Requirement</span></span>| <span data-ttu-id="300f7-873">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-873">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-874">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-874">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-875">1.2</span><span class="sxs-lookup"><span data-stu-id="300f7-875">1.2</span></span>|
|[<span data-ttu-id="300f7-876">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-876">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-877">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="300f7-877">ReadWriteItem</span></span>|
|[<span data-ttu-id="300f7-878">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-878">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-879">Escrever</span><span class="sxs-lookup"><span data-stu-id="300f7-879">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="300f7-880">Retorna:</span><span class="sxs-lookup"><span data-stu-id="300f7-880">Returns:</span></span>

<span data-ttu-id="300f7-881">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="300f7-881">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="300f7-882">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="300f7-882">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="300f7-883">String</span><span class="sxs-lookup"><span data-stu-id="300f7-883">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="300f7-884">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-884">Example</span></span>

```JavaScript
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="300f7-885">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="300f7-885">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="300f7-886">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="300f7-886">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="300f7-p160">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="300f7-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="300f7-890">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="300f7-890">Parameters:</span></span>

|<span data-ttu-id="300f7-891">Nome</span><span class="sxs-lookup"><span data-stu-id="300f7-891">Name</span></span>| <span data-ttu-id="300f7-892">Tipo</span><span class="sxs-lookup"><span data-stu-id="300f7-892">Type</span></span>| <span data-ttu-id="300f7-893">Atributos</span><span class="sxs-lookup"><span data-stu-id="300f7-893">Attributes</span></span>| <span data-ttu-id="300f7-894">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-894">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="300f7-895">function</span><span class="sxs-lookup"><span data-stu-id="300f7-895">function</span></span>||<span data-ttu-id="300f7-896">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="300f7-896">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="300f7-897">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="300f7-897">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="300f7-898">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="300f7-898">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="300f7-899">Objeto</span><span class="sxs-lookup"><span data-stu-id="300f7-899">Object</span></span>| <span data-ttu-id="300f7-900">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-900">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-901">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="300f7-901">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="300f7-902">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="300f7-902">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="300f7-903">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-903">Requirements</span></span>

|<span data-ttu-id="300f7-904">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-904">Requirement</span></span>| <span data-ttu-id="300f7-905">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-906">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-907">1.0</span><span class="sxs-lookup"><span data-stu-id="300f7-907">1.0</span></span>|
|[<span data-ttu-id="300f7-908">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="300f7-909">ReadItem</span></span>|
|[<span data-ttu-id="300f7-910">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-911">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="300f7-911">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-912">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-912">Example</span></span>

<span data-ttu-id="300f7-p163">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="300f7-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="300f7-916">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="300f7-916">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="300f7-917">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="300f7-917">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="300f7-p164">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="300f7-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="300f7-922">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="300f7-922">Parameters:</span></span>

|<span data-ttu-id="300f7-923">Nome</span><span class="sxs-lookup"><span data-stu-id="300f7-923">Name</span></span>| <span data-ttu-id="300f7-924">Tipo</span><span class="sxs-lookup"><span data-stu-id="300f7-924">Type</span></span>| <span data-ttu-id="300f7-925">Atributos</span><span class="sxs-lookup"><span data-stu-id="300f7-925">Attributes</span></span>| <span data-ttu-id="300f7-926">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-926">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="300f7-927">String</span><span class="sxs-lookup"><span data-stu-id="300f7-927">String</span></span>||<span data-ttu-id="300f7-928">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="300f7-928">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="300f7-929">Objeto</span><span class="sxs-lookup"><span data-stu-id="300f7-929">Object</span></span>| <span data-ttu-id="300f7-930">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-930">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-931">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="300f7-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="300f7-932">Objeto</span><span class="sxs-lookup"><span data-stu-id="300f7-932">Object</span></span>| <span data-ttu-id="300f7-933">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-933">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-934">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="300f7-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="300f7-935">function</span><span class="sxs-lookup"><span data-stu-id="300f7-935">function</span></span>| <span data-ttu-id="300f7-936">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-936">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-937">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="300f7-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="300f7-938">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="300f7-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="300f7-939">Erros</span><span class="sxs-lookup"><span data-stu-id="300f7-939">Errors</span></span>

| <span data-ttu-id="300f7-940">Código de erro</span><span class="sxs-lookup"><span data-stu-id="300f7-940">Error code</span></span> | <span data-ttu-id="300f7-941">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="300f7-942">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="300f7-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="300f7-943">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-943">Requirements</span></span>

|<span data-ttu-id="300f7-944">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-944">Requirement</span></span>| <span data-ttu-id="300f7-945">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-946">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-946">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-947">1.1</span><span class="sxs-lookup"><span data-stu-id="300f7-947">1.1</span></span>|
|[<span data-ttu-id="300f7-948">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="300f7-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="300f7-950">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-951">Escrever</span><span class="sxs-lookup"><span data-stu-id="300f7-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-952">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-952">Example</span></span>

<span data-ttu-id="300f7-953">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="300f7-953">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="300f7-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="300f7-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="300f7-955">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="300f7-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="300f7-p165">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="300f7-p165">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="300f7-959">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="300f7-959">Parameters:</span></span>

|<span data-ttu-id="300f7-960">Nome</span><span class="sxs-lookup"><span data-stu-id="300f7-960">Name</span></span>| <span data-ttu-id="300f7-961">Tipo</span><span class="sxs-lookup"><span data-stu-id="300f7-961">Type</span></span>| <span data-ttu-id="300f7-962">Atributos</span><span class="sxs-lookup"><span data-stu-id="300f7-962">Attributes</span></span>| <span data-ttu-id="300f7-963">Descrição</span><span class="sxs-lookup"><span data-stu-id="300f7-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="300f7-964">String</span><span class="sxs-lookup"><span data-stu-id="300f7-964">String</span></span>||<span data-ttu-id="300f7-p166">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="300f7-p166">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="300f7-968">Objeto</span><span class="sxs-lookup"><span data-stu-id="300f7-968">Object</span></span>| <span data-ttu-id="300f7-969">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-969">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-970">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="300f7-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="300f7-971">Objeto</span><span class="sxs-lookup"><span data-stu-id="300f7-971">Object</span></span>| <span data-ttu-id="300f7-972">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-972">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-973">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="300f7-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="300f7-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="300f7-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="300f7-975">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="300f7-975">&lt;optional&gt;</span></span>|<span data-ttu-id="300f7-p167">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="300f7-p167">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="300f7-p168">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="300f7-p168">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="300f7-980">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="300f7-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="300f7-981">function</span><span class="sxs-lookup"><span data-stu-id="300f7-981">function</span></span>||<span data-ttu-id="300f7-982">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="300f7-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="300f7-983">Requisitos</span><span class="sxs-lookup"><span data-stu-id="300f7-983">Requirements</span></span>

|<span data-ttu-id="300f7-984">Requisito</span><span class="sxs-lookup"><span data-stu-id="300f7-984">Requirement</span></span>| <span data-ttu-id="300f7-985">Valor</span><span class="sxs-lookup"><span data-stu-id="300f7-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="300f7-986">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="300f7-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="300f7-987">1.2</span><span class="sxs-lookup"><span data-stu-id="300f7-987">1.2</span></span>|
|[<span data-ttu-id="300f7-988">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="300f7-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="300f7-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="300f7-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="300f7-990">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="300f7-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="300f7-991">Escrever</span><span class="sxs-lookup"><span data-stu-id="300f7-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="300f7-992">Exemplo</span><span class="sxs-lookup"><span data-stu-id="300f7-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
