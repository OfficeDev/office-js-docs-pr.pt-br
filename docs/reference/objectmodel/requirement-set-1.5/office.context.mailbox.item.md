---
title: Office.context.mailbox.item - conjunto de requisitos 1.5
description: ''
ms.date: 01/30/2019
localization_priority: Priority
ms.openlocfilehash: cca0bb4baa15d72a58909ca1417eb52a9bf70a8f
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701838"
---
# <a name="item"></a><span data-ttu-id="98878-102">item</span><span class="sxs-lookup"><span data-stu-id="98878-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="98878-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="98878-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="98878-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="98878-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-106">Requirements</span></span>

|<span data-ttu-id="98878-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-107">Requirement</span></span>| <span data-ttu-id="98878-108">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-110">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-110">1.0</span></span>|
|[<span data-ttu-id="98878-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="98878-112">Restricted</span></span>|
|[<span data-ttu-id="98878-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-114">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="98878-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="98878-115">Members and methods</span></span>

| <span data-ttu-id="98878-116">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-116">Member</span></span> | <span data-ttu-id="98878-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="98878-118">attachments</span><span class="sxs-lookup"><span data-stu-id="98878-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="98878-119">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-119">Member</span></span> |
| [<span data-ttu-id="98878-120">bcc</span><span class="sxs-lookup"><span data-stu-id="98878-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="98878-121">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-121">Member</span></span> |
| [<span data-ttu-id="98878-122">body</span><span class="sxs-lookup"><span data-stu-id="98878-122">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="98878-123">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-123">Member</span></span> |
| [<span data-ttu-id="98878-124">cc</span><span class="sxs-lookup"><span data-stu-id="98878-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="98878-125">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-125">Member</span></span> |
| [<span data-ttu-id="98878-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="98878-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="98878-127">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-127">Member</span></span> |
| [<span data-ttu-id="98878-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="98878-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="98878-129">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-129">Member</span></span> |
| [<span data-ttu-id="98878-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="98878-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="98878-131">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-131">Member</span></span> |
| [<span data-ttu-id="98878-132">end</span><span class="sxs-lookup"><span data-stu-id="98878-132">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="98878-133">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-133">Member</span></span> |
| [<span data-ttu-id="98878-134">from</span><span class="sxs-lookup"><span data-stu-id="98878-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="98878-135">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-135">Member</span></span> |
| [<span data-ttu-id="98878-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="98878-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="98878-137">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-137">Member</span></span> |
| [<span data-ttu-id="98878-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="98878-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="98878-139">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-139">Member</span></span> |
| [<span data-ttu-id="98878-140">itemId</span><span class="sxs-lookup"><span data-stu-id="98878-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="98878-141">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-141">Member</span></span> |
| [<span data-ttu-id="98878-142">itemType</span><span class="sxs-lookup"><span data-stu-id="98878-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="98878-143">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-143">Member</span></span> |
| [<span data-ttu-id="98878-144">location</span><span class="sxs-lookup"><span data-stu-id="98878-144">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="98878-145">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-145">Member</span></span> |
| [<span data-ttu-id="98878-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="98878-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="98878-147">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-147">Member</span></span> |
| [<span data-ttu-id="98878-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="98878-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="98878-149">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-149">Member</span></span> |
| [<span data-ttu-id="98878-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="98878-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="98878-151">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-151">Member</span></span> |
| [<span data-ttu-id="98878-152">organizer</span><span class="sxs-lookup"><span data-stu-id="98878-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="98878-153">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-153">Member</span></span> |
| [<span data-ttu-id="98878-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="98878-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="98878-155">Member</span><span class="sxs-lookup"><span data-stu-id="98878-155">Member</span></span> |
| [<span data-ttu-id="98878-156">sender</span><span class="sxs-lookup"><span data-stu-id="98878-156">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="98878-157">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-157">Member</span></span> |
| [<span data-ttu-id="98878-158">start</span><span class="sxs-lookup"><span data-stu-id="98878-158">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="98878-159">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-159">Member</span></span> |
| [<span data-ttu-id="98878-160">subject</span><span class="sxs-lookup"><span data-stu-id="98878-160">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="98878-161">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-161">Member</span></span> |
| [<span data-ttu-id="98878-162">to</span><span class="sxs-lookup"><span data-stu-id="98878-162">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="98878-163">Membro</span><span class="sxs-lookup"><span data-stu-id="98878-163">Member</span></span> |
| [<span data-ttu-id="98878-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="98878-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="98878-165">Método</span><span class="sxs-lookup"><span data-stu-id="98878-165">Method</span></span> |
| [<span data-ttu-id="98878-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="98878-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="98878-167">Método</span><span class="sxs-lookup"><span data-stu-id="98878-167">Method</span></span> |
| [<span data-ttu-id="98878-168">close</span><span class="sxs-lookup"><span data-stu-id="98878-168">close</span></span>](#close) | <span data-ttu-id="98878-169">Método</span><span class="sxs-lookup"><span data-stu-id="98878-169">Method</span></span> |
| [<span data-ttu-id="98878-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="98878-170">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="98878-171">Método</span><span class="sxs-lookup"><span data-stu-id="98878-171">Method</span></span> |
| [<span data-ttu-id="98878-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="98878-172">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="98878-173">Método</span><span class="sxs-lookup"><span data-stu-id="98878-173">Method</span></span> |
| [<span data-ttu-id="98878-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="98878-174">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="98878-175">Método</span><span class="sxs-lookup"><span data-stu-id="98878-175">Method</span></span> |
| [<span data-ttu-id="98878-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="98878-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="98878-177">Método</span><span class="sxs-lookup"><span data-stu-id="98878-177">Method</span></span> |
| [<span data-ttu-id="98878-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="98878-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="98878-179">Método</span><span class="sxs-lookup"><span data-stu-id="98878-179">Method</span></span> |
| [<span data-ttu-id="98878-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="98878-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="98878-181">Método</span><span class="sxs-lookup"><span data-stu-id="98878-181">Method</span></span> |
| [<span data-ttu-id="98878-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="98878-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="98878-183">Método</span><span class="sxs-lookup"><span data-stu-id="98878-183">Method</span></span> |
| [<span data-ttu-id="98878-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="98878-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="98878-185">Método</span><span class="sxs-lookup"><span data-stu-id="98878-185">Method</span></span> |
| [<span data-ttu-id="98878-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="98878-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="98878-187">Método</span><span class="sxs-lookup"><span data-stu-id="98878-187">Method</span></span> |
| [<span data-ttu-id="98878-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="98878-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="98878-189">Método</span><span class="sxs-lookup"><span data-stu-id="98878-189">Method</span></span> |
| [<span data-ttu-id="98878-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="98878-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="98878-191">Método</span><span class="sxs-lookup"><span data-stu-id="98878-191">Method</span></span> |
| [<span data-ttu-id="98878-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="98878-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="98878-193">Método</span><span class="sxs-lookup"><span data-stu-id="98878-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="98878-194">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-194">Example</span></span>

<span data-ttu-id="98878-195">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="98878-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="98878-196">Membros</span><span class="sxs-lookup"><span data-stu-id="98878-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="98878-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="98878-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="98878-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="98878-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-200">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="98878-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="98878-201">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="98878-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="98878-202">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-202">Type:</span></span>

*   <span data-ttu-id="98878-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="98878-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-204">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-204">Requirements</span></span>

|<span data-ttu-id="98878-205">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-205">Requirement</span></span>| <span data-ttu-id="98878-206">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-207">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-208">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-208">1.0</span></span>|
|[<span data-ttu-id="98878-209">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-209">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-210">ReadItem</span></span>|
|[<span data-ttu-id="98878-211">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-211">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-212">Read</span><span class="sxs-lookup"><span data-stu-id="98878-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-213">Example</span></span>

<span data-ttu-id="98878-214">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="98878-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="98878-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="98878-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="98878-216">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="98878-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="98878-217">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="98878-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-218">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-218">Type:</span></span>

*   [<span data-ttu-id="98878-219">Destinatários</span><span class="sxs-lookup"><span data-stu-id="98878-219">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="98878-220">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-220">Requirements</span></span>

|<span data-ttu-id="98878-221">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-221">Requirement</span></span>| <span data-ttu-id="98878-222">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-223">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-224">1.1</span><span class="sxs-lookup"><span data-stu-id="98878-224">1.1</span></span>|
|[<span data-ttu-id="98878-225">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-225">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-226">ReadItem</span></span>|
|[<span data-ttu-id="98878-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-227">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-228">Escrever</span><span class="sxs-lookup"><span data-stu-id="98878-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-229">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-229">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="98878-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="98878-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="98878-231">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="98878-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-232">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-232">Type:</span></span>

*   [<span data-ttu-id="98878-233">Corpo</span><span class="sxs-lookup"><span data-stu-id="98878-233">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="98878-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-234">Requirements</span></span>

|<span data-ttu-id="98878-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-235">Requirement</span></span>| <span data-ttu-id="98878-236">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-238">1.1</span><span class="sxs-lookup"><span data-stu-id="98878-238">1.1</span></span>|
|[<span data-ttu-id="98878-239">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-240">ReadItem</span></span>|
|[<span data-ttu-id="98878-241">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-242">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-242">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="98878-243">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="98878-243">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="98878-244">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="98878-244">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="98878-245">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="98878-245">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="98878-246">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="98878-246">Read mode</span></span>

<span data-ttu-id="98878-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="98878-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="98878-249">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="98878-249">Compose mode</span></span>

<span data-ttu-id="98878-250">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="98878-250">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-251">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-251">Type:</span></span>

*   <span data-ttu-id="98878-252">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="98878-252">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-253">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-253">Requirements</span></span>

|<span data-ttu-id="98878-254">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-254">Requirement</span></span>| <span data-ttu-id="98878-255">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-256">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-256">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-257">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-257">1.0</span></span>|
|[<span data-ttu-id="98878-258">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-258">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-259">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-259">ReadItem</span></span>|
|[<span data-ttu-id="98878-260">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-260">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-261">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-261">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-262">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-262">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="98878-263">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="98878-263">(nullable) conversationId :String</span></span>

<span data-ttu-id="98878-264">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="98878-264">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="98878-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="98878-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="98878-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="98878-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-269">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-269">Type:</span></span>

*   <span data-ttu-id="98878-270">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="98878-270">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-271">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-271">Requirements</span></span>

|<span data-ttu-id="98878-272">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-272">Requirement</span></span>| <span data-ttu-id="98878-273">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-273">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-274">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-275">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-275">1.0</span></span>|
|[<span data-ttu-id="98878-276">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-277">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-277">ReadItem</span></span>|
|[<span data-ttu-id="98878-278">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-279">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-279">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="98878-280">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="98878-280">dateTimeCreated :Date</span></span>

<span data-ttu-id="98878-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="98878-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-283">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-283">Type:</span></span>

*   <span data-ttu-id="98878-284">Data</span><span class="sxs-lookup"><span data-stu-id="98878-284">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-285">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-285">Requirements</span></span>

|<span data-ttu-id="98878-286">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-286">Requirement</span></span>| <span data-ttu-id="98878-287">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-288">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-289">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-289">1.0</span></span>|
|[<span data-ttu-id="98878-290">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-291">ReadItem</span></span>|
|[<span data-ttu-id="98878-292">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-293">Read</span><span class="sxs-lookup"><span data-stu-id="98878-293">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-294">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-294">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="98878-295">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="98878-295">dateTimeModified :Date</span></span>

<span data-ttu-id="98878-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="98878-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-298">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="98878-298">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-299">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-299">Type:</span></span>

*   <span data-ttu-id="98878-300">Data</span><span class="sxs-lookup"><span data-stu-id="98878-300">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-301">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-301">Requirements</span></span>

|<span data-ttu-id="98878-302">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-302">Requirement</span></span>| <span data-ttu-id="98878-303">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-304">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-305">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-305">1.0</span></span>|
|[<span data-ttu-id="98878-306">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-307">ReadItem</span></span>|
|[<span data-ttu-id="98878-308">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-309">Read</span><span class="sxs-lookup"><span data-stu-id="98878-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-310">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-310">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="98878-311">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="98878-311">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="98878-312">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="98878-312">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="98878-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="98878-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="98878-315">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="98878-315">Read mode</span></span>

<span data-ttu-id="98878-316">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="98878-316">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="98878-317">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="98878-317">Compose mode</span></span>

<span data-ttu-id="98878-318">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="98878-318">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="98878-319">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="98878-319">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-320">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-320">Type:</span></span>

*   <span data-ttu-id="98878-321">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="98878-321">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-322">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-322">Requirements</span></span>

|<span data-ttu-id="98878-323">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-323">Requirement</span></span>| <span data-ttu-id="98878-324">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-324">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-325">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-325">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-326">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-326">1.0</span></span>|
|[<span data-ttu-id="98878-327">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-327">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-328">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-328">ReadItem</span></span>|
|[<span data-ttu-id="98878-329">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-329">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-330">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-330">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-331">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-331">Example</span></span>

<span data-ttu-id="98878-332">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="98878-332">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="98878-333">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="98878-333">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="98878-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="98878-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="98878-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="98878-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-338">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="98878-338">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-339">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-339">Type:</span></span>

*   [<span data-ttu-id="98878-340">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="98878-340">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="98878-341">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-341">Requirements</span></span>

|<span data-ttu-id="98878-342">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-342">Requirement</span></span>| <span data-ttu-id="98878-343">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-344">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-345">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-345">1.0</span></span>|
|[<span data-ttu-id="98878-346">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-346">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-347">ReadItem</span></span>|
|[<span data-ttu-id="98878-348">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-348">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-349">Read</span><span class="sxs-lookup"><span data-stu-id="98878-349">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="98878-350">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="98878-350">internetMessageId :String</span></span>

<span data-ttu-id="98878-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="98878-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-353">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-353">Type:</span></span>

*   <span data-ttu-id="98878-354">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="98878-354">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-355">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-355">Requirements</span></span>

|<span data-ttu-id="98878-356">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-356">Requirement</span></span>| <span data-ttu-id="98878-357">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-358">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-359">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-359">1.0</span></span>|
|[<span data-ttu-id="98878-360">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-361">ReadItem</span></span>|
|[<span data-ttu-id="98878-362">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-363">Read</span><span class="sxs-lookup"><span data-stu-id="98878-363">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-364">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-364">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="98878-365">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="98878-365">itemClass :String</span></span>

<span data-ttu-id="98878-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="98878-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="98878-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="98878-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="98878-370">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-370">Type</span></span> | <span data-ttu-id="98878-371">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-371">Description</span></span> | <span data-ttu-id="98878-372">classe de item</span><span class="sxs-lookup"><span data-stu-id="98878-372">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="98878-373">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="98878-373">Appointment items</span></span> | <span data-ttu-id="98878-374">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="98878-374">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="98878-375">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="98878-375">Message items</span></span> | <span data-ttu-id="98878-376">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="98878-376">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="98878-377">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="98878-377">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-378">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-378">Type:</span></span>

*   <span data-ttu-id="98878-379">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="98878-379">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-380">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-380">Requirements</span></span>

|<span data-ttu-id="98878-381">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-381">Requirement</span></span>| <span data-ttu-id="98878-382">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-382">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-383">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-383">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-384">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-384">1.0</span></span>|
|[<span data-ttu-id="98878-385">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-386">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-386">ReadItem</span></span>|
|[<span data-ttu-id="98878-387">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-388">Read</span><span class="sxs-lookup"><span data-stu-id="98878-388">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-389">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-389">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="98878-390">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="98878-390">(nullable) itemId :String</span></span>

<span data-ttu-id="98878-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="98878-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-393">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="98878-393">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="98878-394">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="98878-394">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="98878-395">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="98878-395">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="98878-396">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="98878-396">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="98878-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="98878-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-399">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-399">Type:</span></span>

*   <span data-ttu-id="98878-400">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="98878-400">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-401">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-401">Requirements</span></span>

|<span data-ttu-id="98878-402">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-402">Requirement</span></span>| <span data-ttu-id="98878-403">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-403">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-404">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-404">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-405">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-405">1.0</span></span>|
|[<span data-ttu-id="98878-406">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-406">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-407">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-407">ReadItem</span></span>|
|[<span data-ttu-id="98878-408">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-408">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-409">Read</span><span class="sxs-lookup"><span data-stu-id="98878-409">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-410">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-410">Example</span></span>

<span data-ttu-id="98878-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="98878-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="98878-413">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="98878-413">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="98878-414">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="98878-414">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="98878-415">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="98878-415">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-416">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-416">Type:</span></span>

*   [<span data-ttu-id="98878-417">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="98878-417">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="98878-418">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-418">Requirements</span></span>

|<span data-ttu-id="98878-419">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-419">Requirement</span></span>| <span data-ttu-id="98878-420">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-421">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-422">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-422">1.0</span></span>|
|[<span data-ttu-id="98878-423">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-424">ReadItem</span></span>|
|[<span data-ttu-id="98878-425">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-426">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-426">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-427">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-427">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="98878-428">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="98878-428">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="98878-429">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="98878-429">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="98878-430">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="98878-430">Read mode</span></span>

<span data-ttu-id="98878-431">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="98878-431">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="98878-432">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="98878-432">Compose mode</span></span>

<span data-ttu-id="98878-433">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="98878-433">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-434">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-434">Type:</span></span>

*   <span data-ttu-id="98878-435">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="98878-435">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-436">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-436">Requirements</span></span>

|<span data-ttu-id="98878-437">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-437">Requirement</span></span>| <span data-ttu-id="98878-438">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-438">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-439">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-440">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-440">1.0</span></span>|
|[<span data-ttu-id="98878-441">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-441">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-442">ReadItem</span></span>|
|[<span data-ttu-id="98878-443">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-443">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-444">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-444">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-445">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-445">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="98878-446">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="98878-446">normalizedSubject :String</span></span>

<span data-ttu-id="98878-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="98878-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="98878-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject).</span><span class="sxs-lookup"><span data-stu-id="98878-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-451">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-451">Type:</span></span>

*   <span data-ttu-id="98878-452">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="98878-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-453">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-453">Requirements</span></span>

|<span data-ttu-id="98878-454">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-454">Requirement</span></span>| <span data-ttu-id="98878-455">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-456">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-457">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-457">1.0</span></span>|
|[<span data-ttu-id="98878-458">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-458">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-459">ReadItem</span></span>|
|[<span data-ttu-id="98878-460">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-460">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-461">Read</span><span class="sxs-lookup"><span data-stu-id="98878-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-462">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-462">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="98878-463">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="98878-463">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="98878-464">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="98878-464">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-465">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-465">Type:</span></span>

*   [<span data-ttu-id="98878-466">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="98878-466">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="98878-467">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-467">Requirements</span></span>

|<span data-ttu-id="98878-468">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-468">Requirement</span></span>| <span data-ttu-id="98878-469">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-470">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-471">1.3</span><span class="sxs-lookup"><span data-stu-id="98878-471">1.3</span></span>|
|[<span data-ttu-id="98878-472">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-472">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-473">ReadItem</span></span>|
|[<span data-ttu-id="98878-474">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-474">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-475">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-475">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="98878-476">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="98878-476">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="98878-477">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="98878-477">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="98878-478">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="98878-478">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="98878-479">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="98878-479">Read mode</span></span>

<span data-ttu-id="98878-480">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="98878-480">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="98878-481">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="98878-481">Compose mode</span></span>

<span data-ttu-id="98878-482">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="98878-482">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-483">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-483">Type:</span></span>

*   <span data-ttu-id="98878-484">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="98878-484">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-485">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-485">Requirements</span></span>

|<span data-ttu-id="98878-486">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-486">Requirement</span></span>| <span data-ttu-id="98878-487">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-488">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-489">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-489">1.0</span></span>|
|[<span data-ttu-id="98878-490">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-491">ReadItem</span></span>|
|[<span data-ttu-id="98878-492">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-493">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-493">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-494">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-494">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="98878-495">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="98878-495">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="98878-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="98878-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-498">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-498">Type:</span></span>

*   [<span data-ttu-id="98878-499">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="98878-499">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="98878-500">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-500">Requirements</span></span>

|<span data-ttu-id="98878-501">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-501">Requirement</span></span>| <span data-ttu-id="98878-502">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-503">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-504">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-504">1.0</span></span>|
|[<span data-ttu-id="98878-505">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-506">ReadItem</span></span>|
|[<span data-ttu-id="98878-507">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-508">Read</span><span class="sxs-lookup"><span data-stu-id="98878-508">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-509">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-509">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="98878-510">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="98878-510">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="98878-511">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="98878-511">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="98878-512">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="98878-512">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="98878-513">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="98878-513">Read mode</span></span>

<span data-ttu-id="98878-514">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="98878-514">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="98878-515">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="98878-515">Compose mode</span></span>

<span data-ttu-id="98878-516">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="98878-516">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-517">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-517">Type:</span></span>

*   <span data-ttu-id="98878-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="98878-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-519">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-519">Requirements</span></span>

|<span data-ttu-id="98878-520">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-520">Requirement</span></span>| <span data-ttu-id="98878-521">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-522">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-523">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-523">1.0</span></span>|
|[<span data-ttu-id="98878-524">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-525">ReadItem</span></span>|
|[<span data-ttu-id="98878-526">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-526">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-527">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-527">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-528">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-528">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="98878-529">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="98878-529">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="98878-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="98878-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="98878-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="98878-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-534">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="98878-534">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-535">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-535">Type:</span></span>

*   [<span data-ttu-id="98878-536">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="98878-536">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="98878-537">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-537">Requirements</span></span>

|<span data-ttu-id="98878-538">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-538">Requirement</span></span>| <span data-ttu-id="98878-539">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-540">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-541">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-541">1.0</span></span>|
|[<span data-ttu-id="98878-542">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-542">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-543">ReadItem</span></span>|
|[<span data-ttu-id="98878-544">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-544">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-545">Read</span><span class="sxs-lookup"><span data-stu-id="98878-545">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-546">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-546">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="98878-547">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="98878-547">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="98878-548">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="98878-548">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="98878-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="98878-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="98878-551">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="98878-551">Read mode</span></span>

<span data-ttu-id="98878-552">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="98878-552">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="98878-553">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="98878-553">Compose mode</span></span>

<span data-ttu-id="98878-554">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="98878-554">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="98878-555">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="98878-555">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-556">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-556">Type:</span></span>

*   <span data-ttu-id="98878-557">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="98878-557">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-558">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-558">Requirements</span></span>

|<span data-ttu-id="98878-559">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-559">Requirement</span></span>| <span data-ttu-id="98878-560">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-561">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-562">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-562">1.0</span></span>|
|[<span data-ttu-id="98878-563">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-564">ReadItem</span></span>|
|[<span data-ttu-id="98878-565">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-566">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-567">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-567">Example</span></span>

<span data-ttu-id="98878-568">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="98878-568">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="98878-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="98878-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="98878-570">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="98878-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="98878-571">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="98878-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="98878-572">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="98878-572">Read mode</span></span>

<span data-ttu-id="98878-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="98878-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="98878-575">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="98878-575">Compose mode</span></span>

<span data-ttu-id="98878-576">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="98878-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="98878-577">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-577">Type:</span></span>

*   <span data-ttu-id="98878-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="98878-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-579">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-579">Requirements</span></span>

|<span data-ttu-id="98878-580">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-580">Requirement</span></span>| <span data-ttu-id="98878-581">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-582">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-583">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-583">1.0</span></span>|
|[<span data-ttu-id="98878-584">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-585">ReadItem</span></span>|
|[<span data-ttu-id="98878-586">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-587">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-587">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="98878-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="98878-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="98878-589">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="98878-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="98878-590">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="98878-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="98878-591">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="98878-591">Read mode</span></span>

<span data-ttu-id="98878-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="98878-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="98878-594">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="98878-594">Compose mode</span></span>

<span data-ttu-id="98878-595">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="98878-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="98878-596">Tipo:</span><span class="sxs-lookup"><span data-stu-id="98878-596">Type:</span></span>

*   <span data-ttu-id="98878-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="98878-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-598">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-598">Requirements</span></span>

|<span data-ttu-id="98878-599">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-599">Requirement</span></span>| <span data-ttu-id="98878-600">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-601">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-602">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-602">1.0</span></span>|
|[<span data-ttu-id="98878-603">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-603">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-604">ReadItem</span></span>|
|[<span data-ttu-id="98878-605">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-605">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-606">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-606">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-607">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-607">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="98878-608">Métodos</span><span class="sxs-lookup"><span data-stu-id="98878-608">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="98878-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="98878-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="98878-610">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="98878-610">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="98878-611">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="98878-611">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="98878-612">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="98878-612">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="98878-613">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="98878-613">Parameters:</span></span>

|<span data-ttu-id="98878-614">Nome</span><span class="sxs-lookup"><span data-stu-id="98878-614">Name</span></span>| <span data-ttu-id="98878-615">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-615">Type</span></span>| <span data-ttu-id="98878-616">Atributos</span><span class="sxs-lookup"><span data-stu-id="98878-616">Attributes</span></span>| <span data-ttu-id="98878-617">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-617">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="98878-618">String</span><span class="sxs-lookup"><span data-stu-id="98878-618">String</span></span>||<span data-ttu-id="98878-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="98878-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="98878-621">String</span><span class="sxs-lookup"><span data-stu-id="98878-621">String</span></span>||<span data-ttu-id="98878-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="98878-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="98878-624">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-624">Object</span></span>| <span data-ttu-id="98878-625">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-625">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-626">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="98878-626">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="98878-627">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-627">Object</span></span> | <span data-ttu-id="98878-628">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-628">&lt;optional&gt;</span></span> | <span data-ttu-id="98878-629">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="98878-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="98878-630">Booliano</span><span class="sxs-lookup"><span data-stu-id="98878-630">Boolean</span></span> | <span data-ttu-id="98878-631">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-631">&lt;optional&gt;</span></span> | <span data-ttu-id="98878-632">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="98878-632">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="98878-633">function</span><span class="sxs-lookup"><span data-stu-id="98878-633">function</span></span>| <span data-ttu-id="98878-634">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-634">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-635">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="98878-635">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="98878-636">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="98878-636">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="98878-637">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="98878-637">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="98878-638">Erros</span><span class="sxs-lookup"><span data-stu-id="98878-638">Errors</span></span>

| <span data-ttu-id="98878-639">Código de erro</span><span class="sxs-lookup"><span data-stu-id="98878-639">Error code</span></span> | <span data-ttu-id="98878-640">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-640">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="98878-641">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="98878-641">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="98878-642">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="98878-642">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="98878-643">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="98878-643">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="98878-644">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-644">Requirements</span></span>

|<span data-ttu-id="98878-645">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-645">Requirement</span></span>| <span data-ttu-id="98878-646">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-646">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-647">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-647">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-648">1.1</span><span class="sxs-lookup"><span data-stu-id="98878-648">1.1</span></span>|
|[<span data-ttu-id="98878-649">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-649">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-650">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="98878-650">ReadWriteItem</span></span>|
|[<span data-ttu-id="98878-651">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-651">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-652">Escrever</span><span class="sxs-lookup"><span data-stu-id="98878-652">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="98878-653">Exemplos</span><span class="sxs-lookup"><span data-stu-id="98878-653">Examples</span></span>

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

<span data-ttu-id="98878-654">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="98878-654">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="98878-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="98878-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="98878-656">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="98878-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="98878-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="98878-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="98878-660">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="98878-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="98878-661">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="98878-661">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="98878-662">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="98878-662">Parameters:</span></span>

|<span data-ttu-id="98878-663">Nome</span><span class="sxs-lookup"><span data-stu-id="98878-663">Name</span></span>| <span data-ttu-id="98878-664">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-664">Type</span></span>| <span data-ttu-id="98878-665">Atributos</span><span class="sxs-lookup"><span data-stu-id="98878-665">Attributes</span></span>| <span data-ttu-id="98878-666">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="98878-667">String</span><span class="sxs-lookup"><span data-stu-id="98878-667">String</span></span>||<span data-ttu-id="98878-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="98878-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="98878-670">String</span><span class="sxs-lookup"><span data-stu-id="98878-670">String</span></span>||<span data-ttu-id="98878-p136">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="98878-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="98878-673">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-673">Object</span></span>| <span data-ttu-id="98878-674">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-674">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-675">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="98878-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="98878-676">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-676">Object</span></span>| <span data-ttu-id="98878-677">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-677">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-678">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="98878-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="98878-679">function</span><span class="sxs-lookup"><span data-stu-id="98878-679">function</span></span>| <span data-ttu-id="98878-680">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-680">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-681">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="98878-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="98878-682">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="98878-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="98878-683">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="98878-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="98878-684">Erros</span><span class="sxs-lookup"><span data-stu-id="98878-684">Errors</span></span>

| <span data-ttu-id="98878-685">Código de erro</span><span class="sxs-lookup"><span data-stu-id="98878-685">Error code</span></span> | <span data-ttu-id="98878-686">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="98878-687">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="98878-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="98878-688">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-688">Requirements</span></span>

|<span data-ttu-id="98878-689">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-689">Requirement</span></span>| <span data-ttu-id="98878-690">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-691">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-692">1.1</span><span class="sxs-lookup"><span data-stu-id="98878-692">1.1</span></span>|
|[<span data-ttu-id="98878-693">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-693">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="98878-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="98878-695">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-695">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-696">Escrever</span><span class="sxs-lookup"><span data-stu-id="98878-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-697">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-697">Example</span></span>

<span data-ttu-id="98878-698">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="98878-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="98878-699">close()</span><span class="sxs-lookup"><span data-stu-id="98878-699">close()</span></span>

<span data-ttu-id="98878-700">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="98878-700">Closes the current item that is being composed.</span></span>

<span data-ttu-id="98878-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="98878-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-703">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="98878-703">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="98878-704">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="98878-704">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-705">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-705">Requirements</span></span>

|<span data-ttu-id="98878-706">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-706">Requirement</span></span>| <span data-ttu-id="98878-707">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-707">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-708">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-708">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-709">1.3</span><span class="sxs-lookup"><span data-stu-id="98878-709">1.3</span></span>|
|[<span data-ttu-id="98878-710">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-710">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-711">Restrito</span><span class="sxs-lookup"><span data-stu-id="98878-711">Restricted</span></span>|
|[<span data-ttu-id="98878-712">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-712">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-713">Escrever</span><span class="sxs-lookup"><span data-stu-id="98878-713">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="98878-714">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="98878-714">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="98878-715">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="98878-715">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-716">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="98878-716">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="98878-717">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="98878-717">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="98878-718">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="98878-718">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="98878-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="98878-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="98878-722">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="98878-722">Parameters:</span></span>

| <span data-ttu-id="98878-723">Nome</span><span class="sxs-lookup"><span data-stu-id="98878-723">Name</span></span> | <span data-ttu-id="98878-724">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-724">Type</span></span> | <span data-ttu-id="98878-725">Atributos</span><span class="sxs-lookup"><span data-stu-id="98878-725">Attributes</span></span> | <span data-ttu-id="98878-726">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-726">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="98878-727">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="98878-727">String &#124; Object</span></span>| |<span data-ttu-id="98878-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="98878-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="98878-730">**OU**</span><span class="sxs-lookup"><span data-stu-id="98878-730">**OR**</span></span><br/><span data-ttu-id="98878-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="98878-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="98878-733">String</span><span class="sxs-lookup"><span data-stu-id="98878-733">String</span></span> | <span data-ttu-id="98878-734">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-734">&lt;optional&gt;</span></span> | <span data-ttu-id="98878-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="98878-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="98878-737">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-737">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="98878-738">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-738">&lt;optional&gt;</span></span> | <span data-ttu-id="98878-739">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="98878-739">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="98878-740">String</span><span class="sxs-lookup"><span data-stu-id="98878-740">String</span></span> | | <span data-ttu-id="98878-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="98878-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="98878-743">String</span><span class="sxs-lookup"><span data-stu-id="98878-743">String</span></span> | | <span data-ttu-id="98878-744">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="98878-744">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="98878-745">String</span><span class="sxs-lookup"><span data-stu-id="98878-745">String</span></span> | | <span data-ttu-id="98878-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="98878-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="98878-748">Booliano</span><span class="sxs-lookup"><span data-stu-id="98878-748">Boolean</span></span> | | <span data-ttu-id="98878-p144">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="98878-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="98878-751">String</span><span class="sxs-lookup"><span data-stu-id="98878-751">String</span></span> | | <span data-ttu-id="98878-p145">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="98878-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="98878-755">function</span><span class="sxs-lookup"><span data-stu-id="98878-755">function</span></span> | <span data-ttu-id="98878-756">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-756">&lt;optional&gt;</span></span> | <span data-ttu-id="98878-757">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="98878-757">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="98878-758">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-758">Requirements</span></span>

|<span data-ttu-id="98878-759">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-759">Requirement</span></span>| <span data-ttu-id="98878-760">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-761">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-762">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-762">1.0</span></span>|
|[<span data-ttu-id="98878-763">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-764">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-764">ReadItem</span></span>|
|[<span data-ttu-id="98878-765">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-766">Read</span><span class="sxs-lookup"><span data-stu-id="98878-766">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="98878-767">Exemplos</span><span class="sxs-lookup"><span data-stu-id="98878-767">Examples</span></span>

<span data-ttu-id="98878-768">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="98878-768">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="98878-769">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="98878-769">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="98878-770">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="98878-770">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="98878-771">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="98878-771">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="98878-772">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="98878-772">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="98878-773">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="98878-773">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="98878-774">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="98878-774">displayReplyForm(formData)</span></span>

<span data-ttu-id="98878-775">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="98878-775">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-776">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="98878-776">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="98878-777">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="98878-777">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="98878-778">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="98878-778">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="98878-p146">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="98878-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="98878-782">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="98878-782">Parameters:</span></span>

| <span data-ttu-id="98878-783">Nome</span><span class="sxs-lookup"><span data-stu-id="98878-783">Name</span></span> | <span data-ttu-id="98878-784">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-784">Type</span></span> | <span data-ttu-id="98878-785">Atributos</span><span class="sxs-lookup"><span data-stu-id="98878-785">Attributes</span></span> | <span data-ttu-id="98878-786">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-786">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="98878-787">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="98878-787">String &#124; Object</span></span>| | <span data-ttu-id="98878-p147">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="98878-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="98878-790">**OU**</span><span class="sxs-lookup"><span data-stu-id="98878-790">**OR**</span></span><br/><span data-ttu-id="98878-p148">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="98878-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="98878-793">String</span><span class="sxs-lookup"><span data-stu-id="98878-793">String</span></span> | <span data-ttu-id="98878-794">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-794">&lt;optional&gt;</span></span> | <span data-ttu-id="98878-p149">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="98878-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="98878-797">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-797">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="98878-798">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-798">&lt;optional&gt;</span></span> | <span data-ttu-id="98878-799">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="98878-799">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="98878-800">String</span><span class="sxs-lookup"><span data-stu-id="98878-800">String</span></span> | | <span data-ttu-id="98878-p150">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="98878-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="98878-803">String</span><span class="sxs-lookup"><span data-stu-id="98878-803">String</span></span> | | <span data-ttu-id="98878-804">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="98878-804">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="98878-805">String</span><span class="sxs-lookup"><span data-stu-id="98878-805">String</span></span> | | <span data-ttu-id="98878-p151">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="98878-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="98878-808">Booliano</span><span class="sxs-lookup"><span data-stu-id="98878-808">Boolean</span></span> | | <span data-ttu-id="98878-p152">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="98878-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="98878-811">String</span><span class="sxs-lookup"><span data-stu-id="98878-811">String</span></span> | | <span data-ttu-id="98878-p153">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="98878-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="98878-815">function</span><span class="sxs-lookup"><span data-stu-id="98878-815">function</span></span> | <span data-ttu-id="98878-816">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-816">&lt;optional&gt;</span></span> | <span data-ttu-id="98878-817">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="98878-817">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="98878-818">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-818">Requirements</span></span>

|<span data-ttu-id="98878-819">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-819">Requirement</span></span>| <span data-ttu-id="98878-820">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-820">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-821">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-821">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-822">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-822">1.0</span></span>|
|[<span data-ttu-id="98878-823">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-823">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-824">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-824">ReadItem</span></span>|
|[<span data-ttu-id="98878-825">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-825">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-826">Read</span><span class="sxs-lookup"><span data-stu-id="98878-826">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="98878-827">Exemplos</span><span class="sxs-lookup"><span data-stu-id="98878-827">Examples</span></span>

<span data-ttu-id="98878-828">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="98878-828">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="98878-829">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="98878-829">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="98878-830">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="98878-830">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="98878-831">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="98878-831">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="98878-832">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="98878-832">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="98878-833">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="98878-833">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="98878-834">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="98878-834">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="98878-835">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="98878-835">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-836">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="98878-836">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-837">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-837">Requirements</span></span>

|<span data-ttu-id="98878-838">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-838">Requirement</span></span>| <span data-ttu-id="98878-839">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-840">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-841">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-841">1.0</span></span>|
|[<span data-ttu-id="98878-842">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-842">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-843">ReadItem</span></span>|
|[<span data-ttu-id="98878-844">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-844">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-845">Read</span><span class="sxs-lookup"><span data-stu-id="98878-845">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="98878-846">Retorna:</span><span class="sxs-lookup"><span data-stu-id="98878-846">Returns:</span></span>

<span data-ttu-id="98878-847">Tipo: [Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="98878-847">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="98878-848">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-848">Example</span></span>

<span data-ttu-id="98878-849">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="98878-849">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="98878-850">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="98878-850">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="98878-851">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="98878-851">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-852">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="98878-852">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="98878-853">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="98878-853">Parameters:</span></span>

|<span data-ttu-id="98878-854">Nome</span><span class="sxs-lookup"><span data-stu-id="98878-854">Name</span></span>| <span data-ttu-id="98878-855">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-855">Type</span></span>| <span data-ttu-id="98878-856">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-856">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="98878-857">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="98878-857">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="98878-858">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="98878-858">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="98878-859">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-859">Requirements</span></span>

|<span data-ttu-id="98878-860">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-860">Requirement</span></span>| <span data-ttu-id="98878-861">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-861">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-862">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-862">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-863">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-863">1.0</span></span>|
|[<span data-ttu-id="98878-864">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-864">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-865">Restrito</span><span class="sxs-lookup"><span data-stu-id="98878-865">Restricted</span></span>|
|[<span data-ttu-id="98878-866">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-866">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-867">Read</span><span class="sxs-lookup"><span data-stu-id="98878-867">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="98878-868">Retorna:</span><span class="sxs-lookup"><span data-stu-id="98878-868">Returns:</span></span>

<span data-ttu-id="98878-869">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="98878-869">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="98878-870">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="98878-870">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="98878-871">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="98878-871">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="98878-872">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="98878-872">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="98878-873">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="98878-873">Value of `entityType`</span></span> | <span data-ttu-id="98878-874">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="98878-874">Type of objects in returned array</span></span> | <span data-ttu-id="98878-875">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="98878-875">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="98878-876">String</span><span class="sxs-lookup"><span data-stu-id="98878-876">String</span></span> | <span data-ttu-id="98878-877">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="98878-877">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="98878-878">Contato</span><span class="sxs-lookup"><span data-stu-id="98878-878">Contact</span></span> | <span data-ttu-id="98878-879">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="98878-879">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="98878-880">String</span><span class="sxs-lookup"><span data-stu-id="98878-880">String</span></span> | <span data-ttu-id="98878-881">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="98878-881">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="98878-882">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="98878-882">MeetingSuggestion</span></span> | <span data-ttu-id="98878-883">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="98878-883">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="98878-884">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="98878-884">PhoneNumber</span></span> | <span data-ttu-id="98878-885">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="98878-885">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="98878-886">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="98878-886">TaskSuggestion</span></span> | <span data-ttu-id="98878-887">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="98878-887">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="98878-888">String</span><span class="sxs-lookup"><span data-stu-id="98878-888">String</span></span> | <span data-ttu-id="98878-889">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="98878-889">**Restricted**</span></span> |

<span data-ttu-id="98878-890">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="98878-890">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="98878-891">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-891">Example</span></span>

<span data-ttu-id="98878-892">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="98878-892">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="98878-893">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="98878-893">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="98878-894">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="98878-894">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-895">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="98878-895">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="98878-896">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="98878-896">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="98878-897">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="98878-897">Parameters:</span></span>

|<span data-ttu-id="98878-898">Nome</span><span class="sxs-lookup"><span data-stu-id="98878-898">Name</span></span>| <span data-ttu-id="98878-899">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-899">Type</span></span>| <span data-ttu-id="98878-900">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-900">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="98878-901">String</span><span class="sxs-lookup"><span data-stu-id="98878-901">String</span></span>|<span data-ttu-id="98878-902">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="98878-902">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="98878-903">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-903">Requirements</span></span>

|<span data-ttu-id="98878-904">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-904">Requirement</span></span>| <span data-ttu-id="98878-905">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-906">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-907">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-907">1.0</span></span>|
|[<span data-ttu-id="98878-908">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-909">ReadItem</span></span>|
|[<span data-ttu-id="98878-910">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-911">Read</span><span class="sxs-lookup"><span data-stu-id="98878-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="98878-912">Retorna:</span><span class="sxs-lookup"><span data-stu-id="98878-912">Returns:</span></span>

<span data-ttu-id="98878-p155">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="98878-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="98878-915">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="98878-915">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="98878-916">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="98878-916">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="98878-917">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="98878-917">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-918">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="98878-918">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="98878-p156">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="98878-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="98878-922">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="98878-922">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="98878-923">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="98878-923">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="98878-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="98878-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="98878-927">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-927">Requirements</span></span>

|<span data-ttu-id="98878-928">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-928">Requirement</span></span>| <span data-ttu-id="98878-929">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-929">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-930">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-930">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-931">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-931">1.0</span></span>|
|[<span data-ttu-id="98878-932">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-932">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-933">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-933">ReadItem</span></span>|
|[<span data-ttu-id="98878-934">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-934">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-935">Read</span><span class="sxs-lookup"><span data-stu-id="98878-935">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="98878-936">Retorna:</span><span class="sxs-lookup"><span data-stu-id="98878-936">Returns:</span></span>

<span data-ttu-id="98878-p158">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="98878-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="98878-939">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="98878-939">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="98878-940">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-940">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="98878-941">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-941">Example</span></span>

<span data-ttu-id="98878-942">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="98878-942">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="98878-943">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="98878-943">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="98878-944">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="98878-944">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-945">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="98878-945">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="98878-946">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="98878-946">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="98878-p159">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="98878-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="98878-949">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="98878-949">Parameters:</span></span>

|<span data-ttu-id="98878-950">Nome</span><span class="sxs-lookup"><span data-stu-id="98878-950">Name</span></span>| <span data-ttu-id="98878-951">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-951">Type</span></span>| <span data-ttu-id="98878-952">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-952">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="98878-953">String</span><span class="sxs-lookup"><span data-stu-id="98878-953">String</span></span>|<span data-ttu-id="98878-954">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="98878-954">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="98878-955">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-955">Requirements</span></span>

|<span data-ttu-id="98878-956">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-956">Requirement</span></span>| <span data-ttu-id="98878-957">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-957">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-958">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-958">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-959">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-959">1.0</span></span>|
|[<span data-ttu-id="98878-960">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-960">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-961">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-961">ReadItem</span></span>|
|[<span data-ttu-id="98878-962">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-962">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-963">Read</span><span class="sxs-lookup"><span data-stu-id="98878-963">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="98878-964">Retorna:</span><span class="sxs-lookup"><span data-stu-id="98878-964">Returns:</span></span>

<span data-ttu-id="98878-965">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="98878-965">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="98878-966">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="98878-966">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="98878-967">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="98878-967">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="98878-968">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-968">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="98878-969">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="98878-969">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="98878-970">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="98878-970">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="98878-p160">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="98878-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="98878-973">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="98878-973">Parameters:</span></span>

|<span data-ttu-id="98878-974">Nome</span><span class="sxs-lookup"><span data-stu-id="98878-974">Name</span></span>| <span data-ttu-id="98878-975">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-975">Type</span></span>| <span data-ttu-id="98878-976">Atributos</span><span class="sxs-lookup"><span data-stu-id="98878-976">Attributes</span></span>| <span data-ttu-id="98878-977">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-977">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="98878-978">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="98878-978">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="98878-p161">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="98878-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="98878-982">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-982">Object</span></span>| <span data-ttu-id="98878-983">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-983">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-984">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="98878-984">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="98878-985">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-985">Object</span></span>| <span data-ttu-id="98878-986">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-986">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-987">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="98878-987">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="98878-988">function</span><span class="sxs-lookup"><span data-stu-id="98878-988">function</span></span>||<span data-ttu-id="98878-989">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="98878-989">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="98878-990">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="98878-990">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="98878-991">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="98878-991">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="98878-992">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-992">Requirements</span></span>

|<span data-ttu-id="98878-993">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-993">Requirement</span></span>| <span data-ttu-id="98878-994">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-994">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-995">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-995">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-996">1.2</span><span class="sxs-lookup"><span data-stu-id="98878-996">1.2</span></span>|
|[<span data-ttu-id="98878-997">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-997">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-998">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="98878-998">ReadWriteItem</span></span>|
|[<span data-ttu-id="98878-999">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-999">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-1000">Escrever</span><span class="sxs-lookup"><span data-stu-id="98878-1000">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="98878-1001">Retorna:</span><span class="sxs-lookup"><span data-stu-id="98878-1001">Returns:</span></span>

<span data-ttu-id="98878-1002">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="98878-1002">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="98878-1003">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="98878-1003">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="98878-1004">String</span><span class="sxs-lookup"><span data-stu-id="98878-1004">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="98878-1005">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-1005">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="98878-1006">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="98878-1006">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="98878-1007">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="98878-1007">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="98878-p163">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="98878-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="98878-1011">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="98878-1011">Parameters:</span></span>

|<span data-ttu-id="98878-1012">Nome</span><span class="sxs-lookup"><span data-stu-id="98878-1012">Name</span></span>| <span data-ttu-id="98878-1013">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-1013">Type</span></span>| <span data-ttu-id="98878-1014">Atributos</span><span class="sxs-lookup"><span data-stu-id="98878-1014">Attributes</span></span>| <span data-ttu-id="98878-1015">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-1015">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="98878-1016">function</span><span class="sxs-lookup"><span data-stu-id="98878-1016">function</span></span>||<span data-ttu-id="98878-1017">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="98878-1017">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="98878-1018">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="98878-1018">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="98878-1019">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="98878-1019">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="98878-1020">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-1020">Object</span></span>| <span data-ttu-id="98878-1021">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-1021">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-1022">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="98878-1022">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="98878-1023">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="98878-1023">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="98878-1024">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-1024">Requirements</span></span>

|<span data-ttu-id="98878-1025">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-1025">Requirement</span></span>| <span data-ttu-id="98878-1026">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-1026">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-1027">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-1027">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-1028">1.0</span><span class="sxs-lookup"><span data-stu-id="98878-1028">1.0</span></span>|
|[<span data-ttu-id="98878-1029">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-1029">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-1030">ReadItem</span><span class="sxs-lookup"><span data-stu-id="98878-1030">ReadItem</span></span>|
|[<span data-ttu-id="98878-1031">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-1031">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-1032">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="98878-1032">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-1033">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-1033">Example</span></span>

<span data-ttu-id="98878-p166">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="98878-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="98878-1037">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="98878-1037">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="98878-1038">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="98878-1038">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="98878-p167">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="98878-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="98878-1043">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="98878-1043">Parameters:</span></span>

|<span data-ttu-id="98878-1044">Nome</span><span class="sxs-lookup"><span data-stu-id="98878-1044">Name</span></span>| <span data-ttu-id="98878-1045">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-1045">Type</span></span>| <span data-ttu-id="98878-1046">Atributos</span><span class="sxs-lookup"><span data-stu-id="98878-1046">Attributes</span></span>| <span data-ttu-id="98878-1047">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-1047">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="98878-1048">String</span><span class="sxs-lookup"><span data-stu-id="98878-1048">String</span></span>||<span data-ttu-id="98878-1049">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="98878-1049">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="98878-1050">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-1050">Object</span></span>| <span data-ttu-id="98878-1051">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-1052">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="98878-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="98878-1053">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-1053">Object</span></span>| <span data-ttu-id="98878-1054">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-1055">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="98878-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="98878-1056">function</span><span class="sxs-lookup"><span data-stu-id="98878-1056">function</span></span>| <span data-ttu-id="98878-1057">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-1058">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="98878-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="98878-1059">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="98878-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="98878-1060">Erros</span><span class="sxs-lookup"><span data-stu-id="98878-1060">Errors</span></span>

| <span data-ttu-id="98878-1061">Código de erro</span><span class="sxs-lookup"><span data-stu-id="98878-1061">Error code</span></span> | <span data-ttu-id="98878-1062">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="98878-1063">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="98878-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="98878-1064">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-1064">Requirements</span></span>

|<span data-ttu-id="98878-1065">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-1065">Requirement</span></span>| <span data-ttu-id="98878-1066">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-1067">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-1067">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="98878-1068">1.1</span></span>|
|[<span data-ttu-id="98878-1069">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="98878-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="98878-1071">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-1072">Escrever</span><span class="sxs-lookup"><span data-stu-id="98878-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-1073">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-1073">Example</span></span>

<span data-ttu-id="98878-1074">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="98878-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="98878-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="98878-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="98878-1076">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="98878-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="98878-p168">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="98878-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-1080">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="98878-1080">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="98878-1081">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="98878-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="98878-p170">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="98878-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="98878-1085">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="98878-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="98878-1086">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="98878-1086">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="98878-1087">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="98878-1087">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="98878-1088">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="98878-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="98878-1089">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="98878-1089">Parameters:</span></span>

|<span data-ttu-id="98878-1090">Nome</span><span class="sxs-lookup"><span data-stu-id="98878-1090">Name</span></span>| <span data-ttu-id="98878-1091">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-1091">Type</span></span>| <span data-ttu-id="98878-1092">Atributos</span><span class="sxs-lookup"><span data-stu-id="98878-1092">Attributes</span></span>| <span data-ttu-id="98878-1093">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="98878-1094">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-1094">Object</span></span>| <span data-ttu-id="98878-1095">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-1096">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="98878-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="98878-1097">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-1097">Object</span></span>| <span data-ttu-id="98878-1098">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-1099">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="98878-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="98878-1100">function</span><span class="sxs-lookup"><span data-stu-id="98878-1100">function</span></span>||<span data-ttu-id="98878-1101">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="98878-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="98878-1102">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="98878-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="98878-1103">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-1103">Requirements</span></span>

|<span data-ttu-id="98878-1104">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-1104">Requirement</span></span>| <span data-ttu-id="98878-1105">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-1106">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="98878-1107">1.3</span></span>|
|[<span data-ttu-id="98878-1108">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="98878-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="98878-1110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-1111">Escrever</span><span class="sxs-lookup"><span data-stu-id="98878-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="98878-1112">Exemplos</span><span class="sxs-lookup"><span data-stu-id="98878-1112">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="98878-p172">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="98878-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="98878-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="98878-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="98878-1116">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="98878-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="98878-p173">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="98878-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="98878-1120">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="98878-1120">Parameters:</span></span>

|<span data-ttu-id="98878-1121">Nome</span><span class="sxs-lookup"><span data-stu-id="98878-1121">Name</span></span>| <span data-ttu-id="98878-1122">Tipo</span><span class="sxs-lookup"><span data-stu-id="98878-1122">Type</span></span>| <span data-ttu-id="98878-1123">Atributos</span><span class="sxs-lookup"><span data-stu-id="98878-1123">Attributes</span></span>| <span data-ttu-id="98878-1124">Descrição</span><span class="sxs-lookup"><span data-stu-id="98878-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="98878-1125">String</span><span class="sxs-lookup"><span data-stu-id="98878-1125">String</span></span>||<span data-ttu-id="98878-p174">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="98878-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="98878-1129">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-1129">Object</span></span>| <span data-ttu-id="98878-1130">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-1131">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="98878-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="98878-1132">Objeto</span><span class="sxs-lookup"><span data-stu-id="98878-1132">Object</span></span>| <span data-ttu-id="98878-1133">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-1134">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="98878-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="98878-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="98878-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="98878-1136">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="98878-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="98878-p175">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="98878-p175">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="98878-p176">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="98878-p176">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="98878-1141">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="98878-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="98878-1142">function</span><span class="sxs-lookup"><span data-stu-id="98878-1142">function</span></span>||<span data-ttu-id="98878-1143">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="98878-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="98878-1144">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98878-1144">Requirements</span></span>

|<span data-ttu-id="98878-1145">Requisito</span><span class="sxs-lookup"><span data-stu-id="98878-1145">Requirement</span></span>| <span data-ttu-id="98878-1146">Valor</span><span class="sxs-lookup"><span data-stu-id="98878-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="98878-1147">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98878-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98878-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="98878-1148">1.2</span></span>|
|[<span data-ttu-id="98878-1149">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="98878-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="98878-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="98878-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="98878-1151">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98878-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="98878-1152">Escrever</span><span class="sxs-lookup"><span data-stu-id="98878-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="98878-1153">Exemplo</span><span class="sxs-lookup"><span data-stu-id="98878-1153">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
