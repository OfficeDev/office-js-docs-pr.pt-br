---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,6
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: e3221ba9cdb8404784f02f75d4f2253432be4f84
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064632"
---
# <a name="item"></a><span data-ttu-id="8fa4e-102">item</span><span class="sxs-lookup"><span data-stu-id="8fa4e-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="8fa4e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="8fa4e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="8fa4e-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-106">Requirements</span></span>

|<span data-ttu-id="8fa4e-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-107">Requirement</span></span>| <span data-ttu-id="8fa4e-108">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-110">1.0</span></span>|
|[<span data-ttu-id="8fa4e-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-112">Restricted</span></span>|
|[<span data-ttu-id="8fa4e-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8fa4e-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-115">Members and methods</span></span>

| <span data-ttu-id="8fa4e-116">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-116">Member</span></span> | <span data-ttu-id="8fa4e-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8fa4e-118">attachments</span><span class="sxs-lookup"><span data-stu-id="8fa4e-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="8fa4e-119">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-119">Member</span></span> |
| [<span data-ttu-id="8fa4e-120">bcc</span><span class="sxs-lookup"><span data-stu-id="8fa4e-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="8fa4e-121">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-121">Member</span></span> |
| [<span data-ttu-id="8fa4e-122">body</span><span class="sxs-lookup"><span data-stu-id="8fa4e-122">body</span></span>](#body-body) | <span data-ttu-id="8fa4e-123">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-123">Member</span></span> |
| [<span data-ttu-id="8fa4e-124">cc</span><span class="sxs-lookup"><span data-stu-id="8fa4e-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8fa4e-125">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-125">Member</span></span> |
| [<span data-ttu-id="8fa4e-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="8fa4e-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="8fa4e-127">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-127">Member</span></span> |
| [<span data-ttu-id="8fa4e-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="8fa4e-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="8fa4e-129">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-129">Member</span></span> |
| [<span data-ttu-id="8fa4e-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="8fa4e-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="8fa4e-131">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-131">Member</span></span> |
| [<span data-ttu-id="8fa4e-132">end</span><span class="sxs-lookup"><span data-stu-id="8fa4e-132">end</span></span>](#end-datetime) | <span data-ttu-id="8fa4e-133">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-133">Member</span></span> |
| [<span data-ttu-id="8fa4e-134">from</span><span class="sxs-lookup"><span data-stu-id="8fa4e-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="8fa4e-135">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-135">Member</span></span> |
| [<span data-ttu-id="8fa4e-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="8fa4e-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="8fa4e-137">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-137">Member</span></span> |
| [<span data-ttu-id="8fa4e-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="8fa4e-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="8fa4e-139">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-139">Member</span></span> |
| [<span data-ttu-id="8fa4e-140">itemId</span><span class="sxs-lookup"><span data-stu-id="8fa4e-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="8fa4e-141">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-141">Member</span></span> |
| [<span data-ttu-id="8fa4e-142">itemType</span><span class="sxs-lookup"><span data-stu-id="8fa4e-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="8fa4e-143">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-143">Member</span></span> |
| [<span data-ttu-id="8fa4e-144">location</span><span class="sxs-lookup"><span data-stu-id="8fa4e-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="8fa4e-145">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-145">Member</span></span> |
| [<span data-ttu-id="8fa4e-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="8fa4e-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="8fa4e-147">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-147">Member</span></span> |
| [<span data-ttu-id="8fa4e-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="8fa4e-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="8fa4e-149">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-149">Member</span></span> |
| [<span data-ttu-id="8fa4e-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="8fa4e-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8fa4e-151">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-151">Member</span></span> |
| [<span data-ttu-id="8fa4e-152">organizer</span><span class="sxs-lookup"><span data-stu-id="8fa4e-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="8fa4e-153">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-153">Member</span></span> |
| [<span data-ttu-id="8fa4e-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="8fa4e-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8fa4e-155">Member</span><span class="sxs-lookup"><span data-stu-id="8fa4e-155">Member</span></span> |
| [<span data-ttu-id="8fa4e-156">sender</span><span class="sxs-lookup"><span data-stu-id="8fa4e-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="8fa4e-157">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-157">Member</span></span> |
| [<span data-ttu-id="8fa4e-158">start</span><span class="sxs-lookup"><span data-stu-id="8fa4e-158">start</span></span>](#start-datetime) | <span data-ttu-id="8fa4e-159">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-159">Member</span></span> |
| [<span data-ttu-id="8fa4e-160">subject</span><span class="sxs-lookup"><span data-stu-id="8fa4e-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="8fa4e-161">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-161">Member</span></span> |
| [<span data-ttu-id="8fa4e-162">to</span><span class="sxs-lookup"><span data-stu-id="8fa4e-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8fa4e-163">Membro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-163">Member</span></span> |
| [<span data-ttu-id="8fa4e-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8fa4e-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="8fa4e-165">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-165">Method</span></span> |
| [<span data-ttu-id="8fa4e-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8fa4e-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="8fa4e-167">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-167">Method</span></span> |
| [<span data-ttu-id="8fa4e-168">close</span><span class="sxs-lookup"><span data-stu-id="8fa4e-168">close</span></span>](#close) | <span data-ttu-id="8fa4e-169">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-169">Method</span></span> |
| [<span data-ttu-id="8fa4e-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="8fa4e-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="8fa4e-171">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-171">Method</span></span> |
| [<span data-ttu-id="8fa4e-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="8fa4e-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="8fa4e-173">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-173">Method</span></span> |
| [<span data-ttu-id="8fa4e-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="8fa4e-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="8fa4e-175">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-175">Method</span></span> |
| [<span data-ttu-id="8fa4e-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="8fa4e-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="8fa4e-177">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-177">Method</span></span> |
| [<span data-ttu-id="8fa4e-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="8fa4e-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="8fa4e-179">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-179">Method</span></span> |
| [<span data-ttu-id="8fa4e-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8fa4e-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="8fa4e-181">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-181">Method</span></span> |
| [<span data-ttu-id="8fa4e-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="8fa4e-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="8fa4e-183">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-183">Method</span></span> |
| [<span data-ttu-id="8fa4e-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8fa4e-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="8fa4e-185">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-185">Method</span></span> |
| [<span data-ttu-id="8fa4e-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="8fa4e-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="8fa4e-187">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-187">Method</span></span> |
| [<span data-ttu-id="8fa4e-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8fa4e-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="8fa4e-189">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-189">Method</span></span> |
| [<span data-ttu-id="8fa4e-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8fa4e-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="8fa4e-191">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-191">Method</span></span> |
| [<span data-ttu-id="8fa4e-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8fa4e-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="8fa4e-193">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-193">Method</span></span> |
| [<span data-ttu-id="8fa4e-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="8fa4e-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="8fa4e-195">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-195">Method</span></span> |
| [<span data-ttu-id="8fa4e-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8fa4e-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="8fa4e-197">Método</span><span class="sxs-lookup"><span data-stu-id="8fa4e-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="8fa4e-198">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-198">Example</span></span>

<span data-ttu-id="8fa4e-199">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="8fa4e-200">Membros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-16"></a><span data-ttu-id="8fa4e-201">anexos: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="8fa4e-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

<span data-ttu-id="8fa4e-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-204">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8fa4e-205">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-206">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-206">Type</span></span>

*   <span data-ttu-id="8fa4e-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="8fa4e-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-208">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-208">Requirements</span></span>

|<span data-ttu-id="8fa4e-209">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-209">Requirement</span></span>| <span data-ttu-id="8fa4e-210">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-211">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-212">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-212">1.0</span></span>|
|[<span data-ttu-id="8fa4e-213">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-214">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-215">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-216">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-217">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-217">Example</span></span>

<span data-ttu-id="8fa4e-218">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="8fa4e-219">CCO: [destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-220">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8fa4e-221">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-222">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-222">Type</span></span>

*   [<span data-ttu-id="8fa4e-223">Destinatários</span><span class="sxs-lookup"><span data-stu-id="8fa4e-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="8fa4e-224">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-224">Requirements</span></span>

|<span data-ttu-id="8fa4e-225">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-225">Requirement</span></span>| <span data-ttu-id="8fa4e-226">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-227">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-228">1.1</span><span class="sxs-lookup"><span data-stu-id="8fa4e-228">1.1</span></span>|
|[<span data-ttu-id="8fa4e-229">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-230">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-231">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-232">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fa4e-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-233">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-233">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-16"></a><span data-ttu-id="8fa4e-234">corpo: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-235">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-236">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-236">Type</span></span>

*   [<span data-ttu-id="8fa4e-237">Body</span><span class="sxs-lookup"><span data-stu-id="8fa4e-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="8fa4e-238">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-238">Requirements</span></span>

|<span data-ttu-id="8fa4e-239">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-239">Requirement</span></span>| <span data-ttu-id="8fa4e-240">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-241">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-242">1.1</span><span class="sxs-lookup"><span data-stu-id="8fa4e-242">1.1</span></span>|
|[<span data-ttu-id="8fa4e-243">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-244">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-245">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-247">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-247">Example</span></span>

<span data-ttu-id="8fa4e-248">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-248">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="8fa4e-249">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="8fa4e-250">[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.6) [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|CC: Array. <</span><span class="sxs-lookup"><span data-stu-id="8fa4e-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-251">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8fa4e-252">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8fa4e-253">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8fa4e-253">Read mode</span></span>

<span data-ttu-id="8fa4e-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="8fa4e-256">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8fa4e-256">Compose mode</span></span>

<span data-ttu-id="8fa4e-257">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8fa4e-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-258">Type</span></span>

*   <span data-ttu-id="8fa4e-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-260">Requirements</span></span>

|<span data-ttu-id="8fa4e-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-261">Requirement</span></span>| <span data-ttu-id="8fa4e-262">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-264">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-264">1.0</span></span>|
|[<span data-ttu-id="8fa4e-265">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-266">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-267">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-268">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-268">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="8fa4e-269">(Nullable) Conversation: String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="8fa4e-270">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8fa4e-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8fa4e-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-275">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-275">Type</span></span>

*   <span data-ttu-id="8fa4e-276">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-277">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-277">Requirements</span></span>

|<span data-ttu-id="8fa4e-278">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-278">Requirement</span></span>| <span data-ttu-id="8fa4e-279">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-280">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-281">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-281">1.0</span></span>|
|[<span data-ttu-id="8fa4e-282">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-283">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-284">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-285">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-286">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-286">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="8fa4e-287">dateTimeCreated: data</span><span class="sxs-lookup"><span data-stu-id="8fa4e-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="8fa4e-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-290">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-290">Type</span></span>

*   <span data-ttu-id="8fa4e-291">Data</span><span class="sxs-lookup"><span data-stu-id="8fa4e-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-292">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-292">Requirements</span></span>

|<span data-ttu-id="8fa4e-293">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-293">Requirement</span></span>| <span data-ttu-id="8fa4e-294">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-295">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-296">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-296">1.0</span></span>|
|[<span data-ttu-id="8fa4e-297">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-298">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-299">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-300">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-301">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-301">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="8fa4e-302">dateTimeModified: data</span><span class="sxs-lookup"><span data-stu-id="8fa4e-302">dateTimeModified: Date</span></span>

<span data-ttu-id="8fa4e-303">Obtém a data e a hora em que um item foi alterado pela última vez.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-303">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="8fa4e-304">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-304">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-305">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-306">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-306">Type</span></span>

*   <span data-ttu-id="8fa4e-307">Data</span><span class="sxs-lookup"><span data-stu-id="8fa4e-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-308">Requirements</span></span>

|<span data-ttu-id="8fa4e-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-309">Requirement</span></span>| <span data-ttu-id="8fa4e-310">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-312">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-312">1.0</span></span>|
|[<span data-ttu-id="8fa4e-313">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-314">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-315">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-316">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-317">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-317">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="8fa4e-318">fim: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-319">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8fa4e-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8fa4e-322">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8fa4e-322">Read mode</span></span>

<span data-ttu-id="8fa4e-323">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-323">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="8fa4e-324">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8fa4e-324">Compose mode</span></span>

<span data-ttu-id="8fa4e-325">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8fa4e-326">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="8fa4e-327">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="8fa4e-328">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-328">Type</span></span>

*   <span data-ttu-id="8fa4e-329">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-330">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-330">Requirements</span></span>

|<span data-ttu-id="8fa4e-331">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-331">Requirement</span></span>| <span data-ttu-id="8fa4e-332">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-333">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-334">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-334">1.0</span></span>|
|[<span data-ttu-id="8fa4e-335">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-336">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-337">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-338">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-338">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="8fa4e-339">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="8fa4e-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-344">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-345">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-345">Type</span></span>

*   [<span data-ttu-id="8fa4e-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8fa4e-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="example"></a><span data-ttu-id="8fa4e-347">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="8fa4e-348">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-348">Requirements</span></span>

|<span data-ttu-id="8fa4e-349">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-349">Requirement</span></span>| <span data-ttu-id="8fa4e-350">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-351">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-352">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-352">1.0</span></span>|
|[<span data-ttu-id="8fa4e-353">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-354">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-355">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-356">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-356">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="8fa4e-357">internetMessageId: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8fa4e-357">internetMessageId: String</span></span>

<span data-ttu-id="8fa4e-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-360">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-360">Type</span></span>

*   <span data-ttu-id="8fa4e-361">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-362">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-362">Requirements</span></span>

|<span data-ttu-id="8fa4e-363">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-363">Requirement</span></span>| <span data-ttu-id="8fa4e-364">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-365">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-366">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-366">1.0</span></span>|
|[<span data-ttu-id="8fa4e-367">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-368">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-369">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-370">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-371">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-371">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="8fa4e-372">doclass: String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-372">itemClass: String</span></span>

<span data-ttu-id="8fa4e-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8fa4e-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="8fa4e-377">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-377">Type</span></span> | <span data-ttu-id="8fa4e-378">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-378">Description</span></span> | <span data-ttu-id="8fa4e-379">classe de item</span><span class="sxs-lookup"><span data-stu-id="8fa4e-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="8fa4e-380">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="8fa4e-380">Appointment items</span></span> | <span data-ttu-id="8fa4e-381">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="8fa4e-382">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-382">Message items</span></span> | <span data-ttu-id="8fa4e-383">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="8fa4e-384">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-385">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-385">Type</span></span>

*   <span data-ttu-id="8fa4e-386">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-387">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-387">Requirements</span></span>

|<span data-ttu-id="8fa4e-388">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-388">Requirement</span></span>| <span data-ttu-id="8fa4e-389">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-390">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-391">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-391">1.0</span></span>|
|[<span data-ttu-id="8fa4e-392">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-393">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-394">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-395">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-396">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-396">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8fa4e-397">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-397">(nullable) itemId: String</span></span>

<span data-ttu-id="8fa4e-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-400">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8fa4e-401">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8fa4e-402">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8fa4e-403">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="8fa4e-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-406">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-406">Type</span></span>

*   <span data-ttu-id="8fa4e-407">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-408">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-408">Requirements</span></span>

|<span data-ttu-id="8fa4e-409">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-409">Requirement</span></span>| <span data-ttu-id="8fa4e-410">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-411">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-412">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-412">1.0</span></span>|
|[<span data-ttu-id="8fa4e-413">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-414">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-415">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-416">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-417">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-417">Example</span></span>

<span data-ttu-id="8fa4e-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-16"></a><span data-ttu-id="8fa4e-420">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-420">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-421">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8fa4e-422">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-423">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-423">Type</span></span>

*   [<span data-ttu-id="8fa4e-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8fa4e-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="8fa4e-425">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-425">Requirements</span></span>

|<span data-ttu-id="8fa4e-426">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-426">Requirement</span></span>| <span data-ttu-id="8fa4e-427">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-428">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-429">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-429">1.0</span></span>|
|[<span data-ttu-id="8fa4e-430">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-431">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-432">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-433">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-434">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-434">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-16"></a><span data-ttu-id="8fa4e-435">local: cadeia de caracteres | [Local](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-435">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-436">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8fa4e-437">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8fa4e-437">Read mode</span></span>

<span data-ttu-id="8fa4e-438">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="8fa4e-439">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8fa4e-439">Compose mode</span></span>

<span data-ttu-id="8fa4e-440">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8fa4e-441">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-441">Type</span></span>

*   <span data-ttu-id="8fa4e-442">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-442">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-443">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-443">Requirements</span></span>

|<span data-ttu-id="8fa4e-444">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-444">Requirement</span></span>| <span data-ttu-id="8fa4e-445">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-446">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-447">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-447">1.0</span></span>|
|[<span data-ttu-id="8fa4e-448">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-449">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-450">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-451">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-451">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8fa4e-452">normalizedSubject: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8fa4e-452">normalizedSubject: String</span></span>

<span data-ttu-id="8fa4e-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8fa4e-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-457">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-457">Type</span></span>

*   <span data-ttu-id="8fa4e-458">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-459">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-459">Requirements</span></span>

|<span data-ttu-id="8fa4e-460">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-460">Requirement</span></span>| <span data-ttu-id="8fa4e-461">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-462">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-463">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-463">1.0</span></span>|
|[<span data-ttu-id="8fa4e-464">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-465">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-466">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-467">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-468">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-468">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-16"></a><span data-ttu-id="8fa4e-469">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-469">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-470">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-471">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-471">Type</span></span>

*   [<span data-ttu-id="8fa4e-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="8fa4e-472">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="8fa4e-473">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-473">Requirements</span></span>

|<span data-ttu-id="8fa4e-474">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-474">Requirement</span></span>| <span data-ttu-id="8fa4e-475">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-476">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-477">1.3</span><span class="sxs-lookup"><span data-stu-id="8fa4e-477">1.3</span></span>|
|[<span data-ttu-id="8fa4e-478">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-479">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-480">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-481">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-482">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-482">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="8fa4e-483">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.6) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="8fa4e-483">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-484">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8fa4e-485">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8fa4e-486">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8fa4e-486">Read mode</span></span>

<span data-ttu-id="8fa4e-487">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="8fa4e-488">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8fa4e-488">Compose mode</span></span>

<span data-ttu-id="8fa4e-489">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8fa4e-490">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-490">Type</span></span>

*   <span data-ttu-id="8fa4e-491">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-491">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-492">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-492">Requirements</span></span>

|<span data-ttu-id="8fa4e-493">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-493">Requirement</span></span>| <span data-ttu-id="8fa4e-494">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-495">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-496">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-496">1.0</span></span>|
|[<span data-ttu-id="8fa4e-497">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-498">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-499">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-500">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-500">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="8fa4e-501">organizador: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-501">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-p124">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-504">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-504">Type</span></span>

*   [<span data-ttu-id="8fa4e-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8fa4e-505">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="8fa4e-506">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-506">Requirements</span></span>

|<span data-ttu-id="8fa4e-507">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-507">Requirement</span></span>| <span data-ttu-id="8fa4e-508">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-509">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-510">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-510">1.0</span></span>|
|[<span data-ttu-id="8fa4e-511">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-512">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-513">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-514">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-515">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-515">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="8fa4e-516">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.6) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="8fa4e-516">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-517">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8fa4e-518">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8fa4e-519">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8fa4e-519">Read mode</span></span>

<span data-ttu-id="8fa4e-520">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="8fa4e-521">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8fa4e-521">Compose mode</span></span>

<span data-ttu-id="8fa4e-522">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="8fa4e-523">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-523">Type</span></span>

*   <span data-ttu-id="8fa4e-524">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-524">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-525">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-525">Requirements</span></span>

|<span data-ttu-id="8fa4e-526">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-526">Requirement</span></span>| <span data-ttu-id="8fa4e-527">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-528">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-529">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-529">1.0</span></span>|
|[<span data-ttu-id="8fa4e-530">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-531">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-532">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-533">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-533">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="8fa4e-534">remetente: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-534">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8fa4e-p127">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-539">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8fa4e-540">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-540">Type</span></span>

*   [<span data-ttu-id="8fa4e-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8fa4e-541">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="8fa4e-542">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-542">Requirements</span></span>

|<span data-ttu-id="8fa4e-543">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-543">Requirement</span></span>| <span data-ttu-id="8fa4e-544">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-545">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-546">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-546">1.0</span></span>|
|[<span data-ttu-id="8fa4e-547">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-548">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-549">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-550">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-551">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-551">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="8fa4e-552">Início: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-552">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-553">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8fa4e-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8fa4e-556">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8fa4e-556">Read mode</span></span>

<span data-ttu-id="8fa4e-557">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-557">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="8fa4e-558">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8fa4e-558">Compose mode</span></span>

<span data-ttu-id="8fa4e-559">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8fa4e-560">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-560">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="8fa4e-561">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="8fa4e-562">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-562">Type</span></span>

*   <span data-ttu-id="8fa4e-563">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-563">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-564">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-564">Requirements</span></span>

|<span data-ttu-id="8fa4e-565">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-565">Requirement</span></span>| <span data-ttu-id="8fa4e-566">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-567">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-568">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-568">1.0</span></span>|
|[<span data-ttu-id="8fa4e-569">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-570">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-571">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-572">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-572">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-16"></a><span data-ttu-id="8fa4e-573">subject: cadeia de caracteres | [Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-573">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-574">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8fa4e-575">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8fa4e-576">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8fa4e-576">Read mode</span></span>

<span data-ttu-id="8fa4e-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="8fa4e-579">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8fa4e-579">Compose mode</span></span>

<span data-ttu-id="8fa4e-580">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="8fa4e-581">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-581">Type</span></span>

*   <span data-ttu-id="8fa4e-582">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-582">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-583">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-583">Requirements</span></span>

|<span data-ttu-id="8fa4e-584">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-584">Requirement</span></span>| <span data-ttu-id="8fa4e-585">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-586">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-587">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-587">1.0</span></span>|
|[<span data-ttu-id="8fa4e-588">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-589">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-590">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-591">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-591">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="8fa4e-592">para: Array. <[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.6) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8fa4e-592">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8fa4e-593">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8fa4e-594">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8fa4e-595">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8fa4e-595">Read mode</span></span>

<span data-ttu-id="8fa4e-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="8fa4e-598">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8fa4e-598">Compose mode</span></span>

<span data-ttu-id="8fa4e-599">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8fa4e-600">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-600">Type</span></span>

*   <span data-ttu-id="8fa4e-601">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-601">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-602">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-602">Requirements</span></span>

|<span data-ttu-id="8fa4e-603">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-603">Requirement</span></span>| <span data-ttu-id="8fa4e-604">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-605">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-606">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-606">1.0</span></span>|
|[<span data-ttu-id="8fa4e-607">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-608">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-609">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-610">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="8fa4e-611">Métodos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-611">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8fa4e-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8fa4e-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8fa4e-613">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8fa4e-614">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8fa4e-615">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8fa4e-616">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-616">Parameters</span></span>

|<span data-ttu-id="8fa4e-617">Nome</span><span class="sxs-lookup"><span data-stu-id="8fa4e-617">Name</span></span>| <span data-ttu-id="8fa4e-618">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-618">Type</span></span>| <span data-ttu-id="8fa4e-619">Atributos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-619">Attributes</span></span>| <span data-ttu-id="8fa4e-620">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="8fa4e-621">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-621">String</span></span>||<span data-ttu-id="8fa4e-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8fa4e-624">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-624">String</span></span>||<span data-ttu-id="8fa4e-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8fa4e-627">Objeto</span><span class="sxs-lookup"><span data-stu-id="8fa4e-627">Object</span></span>| <span data-ttu-id="8fa4e-628">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-628">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-629">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="8fa4e-630">Objeto</span><span class="sxs-lookup"><span data-stu-id="8fa4e-630">Object</span></span> | <span data-ttu-id="8fa4e-631">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-631">&lt;optional&gt;</span></span> | <span data-ttu-id="8fa4e-632">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="8fa4e-633">Booliano</span><span class="sxs-lookup"><span data-stu-id="8fa4e-633">Boolean</span></span> | <span data-ttu-id="8fa4e-634">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-634">&lt;optional&gt;</span></span> | <span data-ttu-id="8fa4e-635">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="8fa4e-636">function</span><span class="sxs-lookup"><span data-stu-id="8fa4e-636">function</span></span>| <span data-ttu-id="8fa4e-637">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-637">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-638">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8fa4e-639">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8fa4e-640">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8fa4e-641">Erros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-641">Errors</span></span>

| <span data-ttu-id="8fa4e-642">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-642">Error code</span></span> | <span data-ttu-id="8fa4e-643">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="8fa4e-644">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="8fa4e-645">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8fa4e-646">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8fa4e-647">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-647">Requirements</span></span>

|<span data-ttu-id="8fa4e-648">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-648">Requirement</span></span>| <span data-ttu-id="8fa4e-649">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-650">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-651">1.1</span><span class="sxs-lookup"><span data-stu-id="8fa4e-651">1.1</span></span>|
|[<span data-ttu-id="8fa4e-652">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="8fa4e-654">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-655">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fa4e-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8fa4e-656">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-656">Examples</span></span>

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

<span data-ttu-id="8fa4e-657">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```javascript
Office.context.mailbox.item.addFileAttachmentAsync(
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
        // Do something here.
      });
  });
```

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8fa4e-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8fa4e-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8fa4e-659">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8fa4e-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8fa4e-663">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8fa4e-664">Se o suplemento do Office estiver em execução no Outlook na Web, o `addItemAttachmentAsync` método poderá anexar itens a itens diferentes do item que você está editando; no entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-664">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8fa4e-665">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-665">Parameters</span></span>

|<span data-ttu-id="8fa4e-666">Nome</span><span class="sxs-lookup"><span data-stu-id="8fa4e-666">Name</span></span>| <span data-ttu-id="8fa4e-667">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-667">Type</span></span>| <span data-ttu-id="8fa4e-668">Atributos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-668">Attributes</span></span>| <span data-ttu-id="8fa4e-669">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="8fa4e-670">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-670">String</span></span>||<span data-ttu-id="8fa4e-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8fa4e-673">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8fa4e-673">String</span></span>||<span data-ttu-id="8fa4e-674">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-674">The subject of the item to be attached.</span></span> <span data-ttu-id="8fa4e-675">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8fa4e-676">Object</span><span class="sxs-lookup"><span data-stu-id="8fa4e-676">Object</span></span>| <span data-ttu-id="8fa4e-677">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-677">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-678">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8fa4e-679">Objeto</span><span class="sxs-lookup"><span data-stu-id="8fa4e-679">Object</span></span>| <span data-ttu-id="8fa4e-680">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-680">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-681">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8fa4e-682">function</span><span class="sxs-lookup"><span data-stu-id="8fa4e-682">function</span></span>| <span data-ttu-id="8fa4e-683">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-683">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-684">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8fa4e-685">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8fa4e-686">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8fa4e-687">Erros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-687">Errors</span></span>

| <span data-ttu-id="8fa4e-688">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-688">Error code</span></span> | <span data-ttu-id="8fa4e-689">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8fa4e-690">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8fa4e-691">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-691">Requirements</span></span>

|<span data-ttu-id="8fa4e-692">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-692">Requirement</span></span>| <span data-ttu-id="8fa4e-693">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-694">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-695">1.1</span><span class="sxs-lookup"><span data-stu-id="8fa4e-695">1.1</span></span>|
|[<span data-ttu-id="8fa4e-696">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="8fa4e-698">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-699">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fa4e-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-700">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-700">Example</span></span>

<span data-ttu-id="8fa4e-701">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="8fa4e-702">close()</span><span class="sxs-lookup"><span data-stu-id="8fa4e-702">close()</span></span>

<span data-ttu-id="8fa4e-703">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="8fa4e-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-706">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="8fa4e-707">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-708">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-708">Requirements</span></span>

|<span data-ttu-id="8fa4e-709">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-709">Requirement</span></span>| <span data-ttu-id="8fa4e-710">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-711">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-712">1.3</span><span class="sxs-lookup"><span data-stu-id="8fa4e-712">1.3</span></span>|
|[<span data-ttu-id="8fa4e-713">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-714">Restrito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-714">Restricted</span></span>|
|[<span data-ttu-id="8fa4e-715">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-716">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fa4e-716">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="8fa4e-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="8fa4e-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="8fa4e-718">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-719">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-719">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8fa4e-720">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-720">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8fa4e-721">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="8fa4e-722">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-722">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="8fa4e-723">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-723">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="8fa4e-724">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-724">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8fa4e-725">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-725">Parameters</span></span>

| <span data-ttu-id="8fa4e-726">Nome</span><span class="sxs-lookup"><span data-stu-id="8fa4e-726">Name</span></span> | <span data-ttu-id="8fa4e-727">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-727">Type</span></span> | <span data-ttu-id="8fa4e-728">Atributos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-728">Attributes</span></span> | <span data-ttu-id="8fa4e-729">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="8fa4e-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8fa4e-730">String &#124; Object</span></span>| |<span data-ttu-id="8fa4e-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8fa4e-733">**OU**</span><span class="sxs-lookup"><span data-stu-id="8fa4e-733">**OR**</span></span><br/><span data-ttu-id="8fa4e-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8fa4e-736">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-736">String</span></span> | <span data-ttu-id="8fa4e-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-737">&lt;optional&gt;</span></span> | <span data-ttu-id="8fa4e-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="8fa4e-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8fa4e-741">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-741">&lt;optional&gt;</span></span> | <span data-ttu-id="8fa4e-742">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="8fa4e-743">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-743">String</span></span> | | <span data-ttu-id="8fa4e-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="8fa4e-746">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8fa4e-746">String</span></span> | | <span data-ttu-id="8fa4e-747">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="8fa4e-748">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-748">String</span></span> | | <span data-ttu-id="8fa4e-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="8fa4e-751">Booliano</span><span class="sxs-lookup"><span data-stu-id="8fa4e-751">Boolean</span></span> | | <span data-ttu-id="8fa4e-p144">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="8fa4e-754">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-754">String</span></span> | | <span data-ttu-id="8fa4e-p145">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="8fa4e-758">function</span><span class="sxs-lookup"><span data-stu-id="8fa4e-758">function</span></span> | <span data-ttu-id="8fa4e-759">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-759">&lt;optional&gt;</span></span> | <span data-ttu-id="8fa4e-760">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8fa4e-761">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-761">Requirements</span></span>

|<span data-ttu-id="8fa4e-762">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-762">Requirement</span></span>| <span data-ttu-id="8fa4e-763">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-764">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-765">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-765">1.0</span></span>|
|[<span data-ttu-id="8fa4e-766">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-767">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-768">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-769">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8fa4e-770">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-770">Examples</span></span>

<span data-ttu-id="8fa4e-771">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8fa4e-772">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-772">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8fa4e-773">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-773">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8fa4e-774">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8fa4e-775">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8fa4e-776">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="8fa4e-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="8fa4e-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="8fa4e-778">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-779">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-779">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8fa4e-780">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-780">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8fa4e-781">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="8fa4e-782">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-782">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="8fa4e-783">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-783">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="8fa4e-784">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-784">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8fa4e-785">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-785">Parameters</span></span>

| <span data-ttu-id="8fa4e-786">Nome</span><span class="sxs-lookup"><span data-stu-id="8fa4e-786">Name</span></span> | <span data-ttu-id="8fa4e-787">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-787">Type</span></span> | <span data-ttu-id="8fa4e-788">Atributos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-788">Attributes</span></span> | <span data-ttu-id="8fa4e-789">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="8fa4e-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8fa4e-790">String &#124; Object</span></span>| | <span data-ttu-id="8fa4e-p147">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8fa4e-793">**OU**</span><span class="sxs-lookup"><span data-stu-id="8fa4e-793">**OR**</span></span><br/><span data-ttu-id="8fa4e-p148">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8fa4e-796">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-796">String</span></span> | <span data-ttu-id="8fa4e-797">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-797">&lt;optional&gt;</span></span> | <span data-ttu-id="8fa4e-p149">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="8fa4e-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8fa4e-801">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-801">&lt;optional&gt;</span></span> | <span data-ttu-id="8fa4e-802">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="8fa4e-803">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-803">String</span></span> | | <span data-ttu-id="8fa4e-p150">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="8fa4e-806">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8fa4e-806">String</span></span> | | <span data-ttu-id="8fa4e-807">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="8fa4e-808">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-808">String</span></span> | | <span data-ttu-id="8fa4e-p151">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="8fa4e-811">Booliano</span><span class="sxs-lookup"><span data-stu-id="8fa4e-811">Boolean</span></span> | | <span data-ttu-id="8fa4e-p152">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="8fa4e-814">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-814">String</span></span> | | <span data-ttu-id="8fa4e-p153">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="8fa4e-818">function</span><span class="sxs-lookup"><span data-stu-id="8fa4e-818">function</span></span> | <span data-ttu-id="8fa4e-819">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-819">&lt;optional&gt;</span></span> | <span data-ttu-id="8fa4e-820">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8fa4e-821">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-821">Requirements</span></span>

|<span data-ttu-id="8fa4e-822">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-822">Requirement</span></span>| <span data-ttu-id="8fa4e-823">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-824">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-825">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-825">1.0</span></span>|
|[<span data-ttu-id="8fa4e-826">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-827">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-828">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-829">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8fa4e-830">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-830">Examples</span></span>

<span data-ttu-id="8fa4e-831">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8fa4e-832">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-832">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8fa4e-833">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-833">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8fa4e-834">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8fa4e-835">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8fa4e-836">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="8fa4e-837">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="8fa4e-837">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="8fa4e-838">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-839">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-839">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-840">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-840">Requirements</span></span>

|<span data-ttu-id="8fa4e-841">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-841">Requirement</span></span>| <span data-ttu-id="8fa4e-842">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-843">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-844">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-844">1.0</span></span>|
|[<span data-ttu-id="8fa4e-845">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-846">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-847">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-848">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8fa4e-849">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8fa4e-849">Returns:</span></span>

<span data-ttu-id="8fa4e-850">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-850">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="8fa4e-851">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-851">Example</span></span>

<span data-ttu-id="8fa4e-852">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-852">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="8fa4e-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="8fa4e-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="8fa4e-854">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-855">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-855">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8fa4e-856">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-856">Parameters</span></span>

|<span data-ttu-id="8fa4e-857">Nome</span><span class="sxs-lookup"><span data-stu-id="8fa4e-857">Name</span></span>| <span data-ttu-id="8fa4e-858">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-858">Type</span></span>| <span data-ttu-id="8fa4e-859">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="8fa4e-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8fa4e-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.6)|<span data-ttu-id="8fa4e-861">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8fa4e-862">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-862">Requirements</span></span>

|<span data-ttu-id="8fa4e-863">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-863">Requirement</span></span>| <span data-ttu-id="8fa4e-864">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-865">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-866">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-866">1.0</span></span>|
|[<span data-ttu-id="8fa4e-867">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-868">Restrito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-868">Restricted</span></span>|
|[<span data-ttu-id="8fa4e-869">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-870">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8fa4e-871">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8fa4e-871">Returns:</span></span>

<span data-ttu-id="8fa4e-872">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8fa4e-873">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="8fa4e-874">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8fa4e-875">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="8fa4e-876">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="8fa4e-876">Value of `entityType`</span></span> | <span data-ttu-id="8fa4e-877">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="8fa4e-877">Type of objects in returned array</span></span> | <span data-ttu-id="8fa4e-878">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="8fa4e-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="8fa4e-879">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-879">String</span></span> | <span data-ttu-id="8fa4e-880">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8fa4e-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="8fa4e-881">Contato</span><span class="sxs-lookup"><span data-stu-id="8fa4e-881">Contact</span></span> | <span data-ttu-id="8fa4e-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8fa4e-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="8fa4e-883">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-883">String</span></span> | <span data-ttu-id="8fa4e-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8fa4e-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="8fa4e-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8fa4e-885">MeetingSuggestion</span></span> | <span data-ttu-id="8fa4e-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8fa4e-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="8fa4e-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8fa4e-887">PhoneNumber</span></span> | <span data-ttu-id="8fa4e-888">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8fa4e-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="8fa4e-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8fa4e-889">TaskSuggestion</span></span> | <span data-ttu-id="8fa4e-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8fa4e-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="8fa4e-891">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-891">String</span></span> | <span data-ttu-id="8fa4e-892">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8fa4e-892">**Restricted**</span></span> |

<span data-ttu-id="8fa4e-893">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="8fa4e-893">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

##### <a name="example"></a><span data-ttu-id="8fa4e-894">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-894">Example</span></span>

<span data-ttu-id="8fa4e-895">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="8fa4e-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="8fa4e-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="8fa4e-897">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-898">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-898">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8fa4e-899">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8fa4e-900">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-900">Parameters</span></span>

|<span data-ttu-id="8fa4e-901">Nome</span><span class="sxs-lookup"><span data-stu-id="8fa4e-901">Name</span></span>| <span data-ttu-id="8fa4e-902">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-902">Type</span></span>| <span data-ttu-id="8fa4e-903">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8fa4e-904">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-904">String</span></span>|<span data-ttu-id="8fa4e-905">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8fa4e-906">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-906">Requirements</span></span>

|<span data-ttu-id="8fa4e-907">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-907">Requirement</span></span>| <span data-ttu-id="8fa4e-908">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-909">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-910">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-910">1.0</span></span>|
|[<span data-ttu-id="8fa4e-911">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-912">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-913">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-914">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8fa4e-915">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8fa4e-915">Returns:</span></span>

<span data-ttu-id="8fa4e-p155">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="8fa4e-918">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="8fa4e-918">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="8fa4e-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8fa4e-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8fa4e-920">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-921">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-921">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8fa4e-p156">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8fa4e-925">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="8fa4e-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8fa4e-926">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8fa4e-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-930">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-930">Requirements</span></span>

|<span data-ttu-id="8fa4e-931">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-931">Requirement</span></span>| <span data-ttu-id="8fa4e-932">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-933">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-934">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-934">1.0</span></span>|
|[<span data-ttu-id="8fa4e-935">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-936">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-937">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-938">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8fa4e-939">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8fa4e-939">Returns:</span></span>

<span data-ttu-id="8fa4e-p158">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="8fa4e-942">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="8fa4e-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8fa4e-943">Objeto</span><span class="sxs-lookup"><span data-stu-id="8fa4e-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8fa4e-944">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-944">Example</span></span>

<span data-ttu-id="8fa4e-945">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8fa4e-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="8fa4e-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8fa4e-947">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-948">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-948">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8fa4e-949">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8fa4e-p159">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8fa4e-952">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-952">Parameters</span></span>

|<span data-ttu-id="8fa4e-953">Nome</span><span class="sxs-lookup"><span data-stu-id="8fa4e-953">Name</span></span>| <span data-ttu-id="8fa4e-954">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-954">Type</span></span>| <span data-ttu-id="8fa4e-955">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8fa4e-956">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-956">String</span></span>|<span data-ttu-id="8fa4e-957">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8fa4e-958">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-958">Requirements</span></span>

|<span data-ttu-id="8fa4e-959">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-959">Requirement</span></span>| <span data-ttu-id="8fa4e-960">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-961">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-962">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-962">1.0</span></span>|
|[<span data-ttu-id="8fa4e-963">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-963">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-964">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-965">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-965">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-966">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8fa4e-967">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8fa4e-967">Returns:</span></span>

<span data-ttu-id="8fa4e-968">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="8fa4e-969">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="8fa4e-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8fa4e-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="8fa4e-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8fa4e-971">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-971">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="8fa4e-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="8fa4e-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="8fa4e-973">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="8fa4e-p160">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8fa4e-976">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-976">Parameters</span></span>

|<span data-ttu-id="8fa4e-977">Nome</span><span class="sxs-lookup"><span data-stu-id="8fa4e-977">Name</span></span>| <span data-ttu-id="8fa4e-978">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-978">Type</span></span>| <span data-ttu-id="8fa4e-979">Atributos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-979">Attributes</span></span>| <span data-ttu-id="8fa4e-980">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="8fa4e-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8fa4e-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="8fa4e-p161">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="8fa4e-985">Objeto</span><span class="sxs-lookup"><span data-stu-id="8fa4e-985">Object</span></span>| <span data-ttu-id="8fa4e-986">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-986">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-987">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8fa4e-988">Object</span><span class="sxs-lookup"><span data-stu-id="8fa4e-988">Object</span></span>| <span data-ttu-id="8fa4e-989">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-989">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-990">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8fa4e-991">function</span><span class="sxs-lookup"><span data-stu-id="8fa4e-991">function</span></span>||<span data-ttu-id="8fa4e-992">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8fa4e-993">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="8fa4e-994">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8fa4e-995">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-995">Requirements</span></span>

|<span data-ttu-id="8fa4e-996">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-996">Requirement</span></span>| <span data-ttu-id="8fa4e-997">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-998">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-999">1.2</span><span class="sxs-lookup"><span data-stu-id="8fa4e-999">1.2</span></span>|
|[<span data-ttu-id="8fa4e-1000">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1000">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="8fa4e-1002">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1002">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-1003">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="8fa4e-1004">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1004">Returns:</span></span>

<span data-ttu-id="8fa4e-1005">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="8fa4e-1006">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8fa4e-1007">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8fa4e-1008">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="8fa4e-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="8fa4e-1010">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1010">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="8fa4e-1011">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1011">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-1012">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1012">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-1013">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1013">Requirements</span></span>

|<span data-ttu-id="8fa4e-1014">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1014">Requirement</span></span>| <span data-ttu-id="8fa4e-1015">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-1016">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1017">1.6</span></span> |
|[<span data-ttu-id="8fa4e-1018">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1018">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1019">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-1020">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1020">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-1021">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8fa4e-1022">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1022">Returns:</span></span>

<span data-ttu-id="8fa4e-1023">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1023">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="8fa4e-1024">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1024">Example</span></span>

<span data-ttu-id="8fa4e-1025">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="8fa4e-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="8fa4e-p164">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-1029">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1029">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8fa4e-p165">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8fa4e-1033">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8fa4e-1034">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8fa4e-p166">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fa4e-1038">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1038">Requirements</span></span>

|<span data-ttu-id="8fa4e-1039">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1039">Requirement</span></span>| <span data-ttu-id="8fa4e-1040">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-1041">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1042">1.6</span></span> |
|[<span data-ttu-id="8fa4e-1043">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1044">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-1045">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-1046">Read</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8fa4e-1047">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1047">Returns:</span></span>

<span data-ttu-id="8fa4e-p167">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="8fa4e-1050">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1050">Example</span></span>

<span data-ttu-id="8fa4e-1051">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8fa4e-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8fa4e-1053">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8fa4e-p168">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8fa4e-1057">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1057">Parameters</span></span>

|<span data-ttu-id="8fa4e-1058">Nome</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1058">Name</span></span>| <span data-ttu-id="8fa4e-1059">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1059">Type</span></span>| <span data-ttu-id="8fa4e-1060">Atributos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1060">Attributes</span></span>| <span data-ttu-id="8fa4e-1061">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8fa4e-1062">function</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1062">function</span></span>||<span data-ttu-id="8fa4e-1063">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8fa4e-1064">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8fa4e-1065">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="8fa4e-1066">Objeto</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1066">Object</span></span>| <span data-ttu-id="8fa4e-1067">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-1068">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="8fa4e-1069">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8fa4e-1070">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1070">Requirements</span></span>

|<span data-ttu-id="8fa4e-1071">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1071">Requirement</span></span>| <span data-ttu-id="8fa4e-1072">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-1073">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1074">1.0</span></span>|
|[<span data-ttu-id="8fa4e-1075">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1075">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1076">ReadItem</span></span>|
|[<span data-ttu-id="8fa4e-1077">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1077">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-1078">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1078">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-1079">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1079">Example</span></span>

<span data-ttu-id="8fa4e-p171">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8fa4e-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8fa4e-1084">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8fa4e-1085">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1085">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="8fa4e-1086">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1086">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="8fa4e-1087">No Outlook na Web e dispositivos móveis, o identificador de anexo é válido somente dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1087">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="8fa4e-1088">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1088">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8fa4e-1089">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1089">Parameters</span></span>

|<span data-ttu-id="8fa4e-1090">Nome</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1090">Name</span></span>| <span data-ttu-id="8fa4e-1091">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1091">Type</span></span>| <span data-ttu-id="8fa4e-1092">Atributos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1092">Attributes</span></span>| <span data-ttu-id="8fa4e-1093">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="8fa4e-1094">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1094">String</span></span>||<span data-ttu-id="8fa4e-1095">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1095">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="8fa4e-1096">Objeto</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1096">Object</span></span>| <span data-ttu-id="8fa4e-1097">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-1098">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1098">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8fa4e-1099">Objeto</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1099">Object</span></span>| <span data-ttu-id="8fa4e-1100">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-1101">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1101">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8fa4e-1102">function</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1102">function</span></span>| <span data-ttu-id="8fa4e-1103">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1103">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-1104">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1104">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8fa4e-1105">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1105">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8fa4e-1106">Erros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1106">Errors</span></span>

| <span data-ttu-id="8fa4e-1107">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1107">Error code</span></span> | <span data-ttu-id="8fa4e-1108">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1108">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="8fa4e-1109">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1109">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8fa4e-1110">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1110">Requirements</span></span>

|<span data-ttu-id="8fa4e-1111">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1111">Requirement</span></span>| <span data-ttu-id="8fa4e-1112">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-1113">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-1114">1.1</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1114">1.1</span></span>|
|[<span data-ttu-id="8fa4e-1115">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-1116">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1116">ReadWriteItem</span></span>|
|[<span data-ttu-id="8fa4e-1117">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-1118">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1118">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-1119">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1119">Example</span></span>

<span data-ttu-id="8fa4e-1120">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1120">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="8fa4e-1121">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1121">saveAsync([options], callback)</span></span>

<span data-ttu-id="8fa4e-1122">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1122">Asynchronously saves an item.</span></span>

<span data-ttu-id="8fa4e-1123">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1123">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="8fa4e-1124">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1124">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="8fa4e-1125">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1125">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-1126">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1126">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="8fa4e-1127">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1127">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="8fa4e-p175">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="8fa4e-1131">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1131">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="8fa4e-1132">O Outlook no Mac não dá suporte à gravação de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1132">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="8fa4e-1133">O `saveAsync` método falha quando chamado de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1133">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="8fa4e-1134">Consulte [não é possível salvar uma reunião como rascunho no Outlook para Mac usando a API do Office js](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1134">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="8fa4e-1135">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8fa4e-1136">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1136">Parameters</span></span>

|<span data-ttu-id="8fa4e-1137">Nome</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1137">Name</span></span>| <span data-ttu-id="8fa4e-1138">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1138">Type</span></span>| <span data-ttu-id="8fa4e-1139">Atributos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1139">Attributes</span></span>| <span data-ttu-id="8fa4e-1140">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="8fa4e-1141">Object</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1141">Object</span></span>| <span data-ttu-id="8fa4e-1142">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-1143">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8fa4e-1144">Objeto</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1144">Object</span></span>| <span data-ttu-id="8fa4e-1145">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-1146">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8fa4e-1147">function</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1147">function</span></span>||<span data-ttu-id="8fa4e-1148">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8fa4e-1149">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1149">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8fa4e-1150">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1150">Requirements</span></span>

|<span data-ttu-id="8fa4e-1151">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1151">Requirement</span></span>| <span data-ttu-id="8fa4e-1152">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-1153">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1154">1.3</span></span>|
|[<span data-ttu-id="8fa4e-1155">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1155">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="8fa4e-1157">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1157">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-1158">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8fa4e-1159">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1159">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="8fa4e-p177">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="8fa4e-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="8fa4e-1163">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="8fa4e-p178">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8fa4e-1167">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1167">Parameters</span></span>

|<span data-ttu-id="8fa4e-1168">Nome</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1168">Name</span></span>| <span data-ttu-id="8fa4e-1169">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1169">Type</span></span>| <span data-ttu-id="8fa4e-1170">Atributos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1170">Attributes</span></span>| <span data-ttu-id="8fa4e-1171">Descrição</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="8fa4e-1172">String</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1172">String</span></span>||<span data-ttu-id="8fa4e-p179">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="8fa4e-1176">Objeto</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1176">Object</span></span>| <span data-ttu-id="8fa4e-1177">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-1178">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8fa4e-1179">Objeto</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1179">Object</span></span>| <span data-ttu-id="8fa4e-1180">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-1181">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="8fa4e-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="8fa4e-1183">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="8fa4e-1184">Se `text`, o estilo atual é aplicado no Outlook na Web e clientes da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1184">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="8fa4e-1185">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1185">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="8fa4e-1186">Se `html` e o campo oferecer suporte a HTML (o assunto não), o estilo atual será aplicado no Outlook na Web e o estilo padrão será aplicado nos clientes da área de trabalho do Outlook.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="8fa4e-1187">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1187">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="8fa4e-1188">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="8fa4e-1189">function</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1189">function</span></span>||<span data-ttu-id="8fa4e-1190">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8fa4e-1191">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1191">Requirements</span></span>

|<span data-ttu-id="8fa4e-1192">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1192">Requirement</span></span>| <span data-ttu-id="8fa4e-1193">Valor</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fa4e-1194">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8fa4e-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1195">1.2</span></span>|
|[<span data-ttu-id="8fa4e-1196">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8fa4e-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="8fa4e-1198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8fa4e-1199">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8fa4e-1200">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fa4e-1200">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
