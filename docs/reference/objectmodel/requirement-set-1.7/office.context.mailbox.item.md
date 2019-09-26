---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,7
description: ''
ms.date: 09/23/2019
localization_priority: Normal
ms.openlocfilehash: 9667ccfea2a27a543ead6df98b24bfa4ca233af4
ms.sourcegitcommit: 3c84fe6302341668c3f9f6dd64e636a97d03023c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/26/2019
ms.locfileid: "37167323"
---
# <a name="item"></a><span data-ttu-id="bd479-102">item</span><span class="sxs-lookup"><span data-stu-id="bd479-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="bd479-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="bd479-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="bd479-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="bd479-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-106">Requirements</span></span>

|<span data-ttu-id="bd479-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-107">Requirement</span></span>|<span data-ttu-id="bd479-108">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-110">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-110">1.0</span></span>|
|[<span data-ttu-id="bd479-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="bd479-112">Restricted</span></span>|
|[<span data-ttu-id="bd479-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bd479-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="bd479-115">Members and methods</span></span>

| <span data-ttu-id="bd479-116">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-116">Member</span></span> | <span data-ttu-id="bd479-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bd479-118">attachments</span><span class="sxs-lookup"><span data-stu-id="bd479-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="bd479-119">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-119">Member</span></span> |
| [<span data-ttu-id="bd479-120">bcc</span><span class="sxs-lookup"><span data-stu-id="bd479-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="bd479-121">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-121">Member</span></span> |
| [<span data-ttu-id="bd479-122">body</span><span class="sxs-lookup"><span data-stu-id="bd479-122">body</span></span>](#body-body) | <span data-ttu-id="bd479-123">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-123">Member</span></span> |
| [<span data-ttu-id="bd479-124">cc</span><span class="sxs-lookup"><span data-stu-id="bd479-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="bd479-125">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-125">Member</span></span> |
| [<span data-ttu-id="bd479-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="bd479-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="bd479-127">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-127">Member</span></span> |
| [<span data-ttu-id="bd479-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="bd479-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="bd479-129">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-129">Member</span></span> |
| [<span data-ttu-id="bd479-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="bd479-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="bd479-131">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-131">Member</span></span> |
| [<span data-ttu-id="bd479-132">end</span><span class="sxs-lookup"><span data-stu-id="bd479-132">end</span></span>](#end-datetime) | <span data-ttu-id="bd479-133">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-133">Member</span></span> |
| [<span data-ttu-id="bd479-134">from</span><span class="sxs-lookup"><span data-stu-id="bd479-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="bd479-135">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-135">Member</span></span> |
| [<span data-ttu-id="bd479-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="bd479-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="bd479-137">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-137">Member</span></span> |
| [<span data-ttu-id="bd479-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="bd479-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="bd479-139">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-139">Member</span></span> |
| [<span data-ttu-id="bd479-140">itemId</span><span class="sxs-lookup"><span data-stu-id="bd479-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="bd479-141">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-141">Member</span></span> |
| [<span data-ttu-id="bd479-142">itemType</span><span class="sxs-lookup"><span data-stu-id="bd479-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="bd479-143">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-143">Member</span></span> |
| [<span data-ttu-id="bd479-144">location</span><span class="sxs-lookup"><span data-stu-id="bd479-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="bd479-145">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-145">Member</span></span> |
| [<span data-ttu-id="bd479-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="bd479-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="bd479-147">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-147">Member</span></span> |
| [<span data-ttu-id="bd479-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="bd479-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="bd479-149">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-149">Member</span></span> |
| [<span data-ttu-id="bd479-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="bd479-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="bd479-151">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-151">Member</span></span> |
| [<span data-ttu-id="bd479-152">organizer</span><span class="sxs-lookup"><span data-stu-id="bd479-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="bd479-153">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-153">Member</span></span> |
| [<span data-ttu-id="bd479-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="bd479-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="bd479-155">Member</span><span class="sxs-lookup"><span data-stu-id="bd479-155">Member</span></span> |
| [<span data-ttu-id="bd479-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="bd479-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="bd479-157">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-157">Member</span></span> |
| [<span data-ttu-id="bd479-158">sender</span><span class="sxs-lookup"><span data-stu-id="bd479-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="bd479-159">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-159">Member</span></span> |
| [<span data-ttu-id="bd479-160">seriesid</span><span class="sxs-lookup"><span data-stu-id="bd479-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="bd479-161">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-161">Member</span></span> |
| [<span data-ttu-id="bd479-162">start</span><span class="sxs-lookup"><span data-stu-id="bd479-162">start</span></span>](#start-datetime) | <span data-ttu-id="bd479-163">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-163">Member</span></span> |
| [<span data-ttu-id="bd479-164">subject</span><span class="sxs-lookup"><span data-stu-id="bd479-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="bd479-165">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-165">Member</span></span> |
| [<span data-ttu-id="bd479-166">to</span><span class="sxs-lookup"><span data-stu-id="bd479-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="bd479-167">Membro</span><span class="sxs-lookup"><span data-stu-id="bd479-167">Member</span></span> |
| [<span data-ttu-id="bd479-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="bd479-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="bd479-169">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-169">Method</span></span> |
| [<span data-ttu-id="bd479-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="bd479-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="bd479-171">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-171">Method</span></span> |
| [<span data-ttu-id="bd479-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="bd479-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="bd479-173">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-173">Method</span></span> |
| [<span data-ttu-id="bd479-174">close</span><span class="sxs-lookup"><span data-stu-id="bd479-174">close</span></span>](#close) | <span data-ttu-id="bd479-175">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-175">Method</span></span> |
| [<span data-ttu-id="bd479-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="bd479-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="bd479-177">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-177">Method</span></span> |
| [<span data-ttu-id="bd479-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="bd479-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="bd479-179">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-179">Method</span></span> |
| [<span data-ttu-id="bd479-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="bd479-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="bd479-181">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-181">Method</span></span> |
| [<span data-ttu-id="bd479-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="bd479-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="bd479-183">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-183">Method</span></span> |
| [<span data-ttu-id="bd479-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="bd479-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="bd479-185">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-185">Method</span></span> |
| [<span data-ttu-id="bd479-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="bd479-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="bd479-187">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-187">Method</span></span> |
| [<span data-ttu-id="bd479-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="bd479-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="bd479-189">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-189">Method</span></span> |
| [<span data-ttu-id="bd479-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="bd479-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="bd479-191">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-191">Method</span></span> |
| [<span data-ttu-id="bd479-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="bd479-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="bd479-193">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-193">Method</span></span> |
| [<span data-ttu-id="bd479-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="bd479-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="bd479-195">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-195">Method</span></span> |
| [<span data-ttu-id="bd479-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="bd479-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="bd479-197">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-197">Method</span></span> |
| [<span data-ttu-id="bd479-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="bd479-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="bd479-199">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-199">Method</span></span> |
| [<span data-ttu-id="bd479-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="bd479-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="bd479-201">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-201">Method</span></span> |
| [<span data-ttu-id="bd479-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="bd479-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="bd479-203">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-203">Method</span></span> |
| [<span data-ttu-id="bd479-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="bd479-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="bd479-205">Método</span><span class="sxs-lookup"><span data-stu-id="bd479-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="bd479-206">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-206">Example</span></span>

<span data-ttu-id="bd479-207">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="bd479-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
};
```

### <a name="members"></a><span data-ttu-id="bd479-208">Membros</span><span class="sxs-lookup"><span data-stu-id="bd479-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-17"></a><span data-ttu-id="bd479-209">anexos: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="bd479-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

<span data-ttu-id="bd479-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="bd479-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-212">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="bd479-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="bd479-213">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="bd479-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-214">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-214">Type</span></span>

*   <span data-ttu-id="bd479-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="bd479-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-216">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-216">Requirements</span></span>

|<span data-ttu-id="bd479-217">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-217">Requirement</span></span>|<span data-ttu-id="bd479-218">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-219">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-220">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-220">1.0</span></span>|
|[<span data-ttu-id="bd479-221">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-222">ReadItem</span></span>|
|[<span data-ttu-id="bd479-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-224">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-225">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-225">Example</span></span>

<span data-ttu-id="bd479-226">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="bd479-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
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

<br>

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="bd479-227">CCO: [destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-228">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="bd479-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="bd479-229">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="bd479-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-230">Type</span></span>

*   [<span data-ttu-id="bd479-231">Destinatários</span><span class="sxs-lookup"><span data-stu-id="bd479-231">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="bd479-232">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-232">Requirements</span></span>

|<span data-ttu-id="bd479-233">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-233">Requirement</span></span>|<span data-ttu-id="bd479-234">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-235">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-236">1.1</span><span class="sxs-lookup"><span data-stu-id="bd479-236">1.1</span></span>|
|[<span data-ttu-id="bd479-237">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-238">ReadItem</span></span>|
|[<span data-ttu-id="bd479-239">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-240">Escrever</span><span class="sxs-lookup"><span data-stu-id="bd479-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-241">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-241">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

<br>

---
---

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-17"></a><span data-ttu-id="bd479-242">corpo: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-242">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-243">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="bd479-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-244">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-244">Type</span></span>

*   [<span data-ttu-id="bd479-245">Body</span><span class="sxs-lookup"><span data-stu-id="bd479-245">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="bd479-246">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-246">Requirements</span></span>

|<span data-ttu-id="bd479-247">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-247">Requirement</span></span>|<span data-ttu-id="bd479-248">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-249">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-250">1.1</span><span class="sxs-lookup"><span data-stu-id="bd479-250">1.1</span></span>|
|[<span data-ttu-id="bd479-251">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-251">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-252">ReadItem</span></span>|
|[<span data-ttu-id="bd479-253">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="bd479-253">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-254">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-254">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-255">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-255">Example</span></span>

<span data-ttu-id="bd479-256">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="bd479-256">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="bd479-257">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-257">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="bd479-258">[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7) [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|CC: Array. <</span><span class="sxs-lookup"><span data-stu-id="bd479-258">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-259">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="bd479-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="bd479-260">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="bd479-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bd479-261">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="bd479-261">Read mode</span></span>

<span data-ttu-id="bd479-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="bd479-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="bd479-264">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="bd479-264">Compose mode</span></span>

<span data-ttu-id="bd479-265">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="bd479-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bd479-266">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-266">Type</span></span>

*   <span data-ttu-id="bd479-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-268">Requirements</span></span>

|<span data-ttu-id="bd479-269">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-269">Requirement</span></span>|<span data-ttu-id="bd479-270">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-271">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-272">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-272">1.0</span></span>|
|[<span data-ttu-id="bd479-273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-274">ReadItem</span></span>|
|[<span data-ttu-id="bd479-275">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="bd479-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-276">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-276">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="bd479-277">(Nullable) Conversation: String</span><span class="sxs-lookup"><span data-stu-id="bd479-277">(nullable) conversationId: String</span></span>

<span data-ttu-id="bd479-278">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="bd479-278">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="bd479-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="bd479-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="bd479-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="bd479-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-283">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-283">Type</span></span>

*   <span data-ttu-id="bd479-284">String</span><span class="sxs-lookup"><span data-stu-id="bd479-284">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-285">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-285">Requirements</span></span>

|<span data-ttu-id="bd479-286">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-286">Requirement</span></span>|<span data-ttu-id="bd479-287">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-288">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-289">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-289">1.0</span></span>|
|[<span data-ttu-id="bd479-290">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-291">ReadItem</span></span>|
|[<span data-ttu-id="bd479-292">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="bd479-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-293">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-294">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-294">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="bd479-295">dateTimeCreated: data</span><span class="sxs-lookup"><span data-stu-id="bd479-295">dateTimeCreated: Date</span></span>

<span data-ttu-id="bd479-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="bd479-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-298">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-298">Type</span></span>

*   <span data-ttu-id="bd479-299">Data</span><span class="sxs-lookup"><span data-stu-id="bd479-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-300">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-300">Requirements</span></span>

|<span data-ttu-id="bd479-301">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-301">Requirement</span></span>|<span data-ttu-id="bd479-302">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-303">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-304">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-304">1.0</span></span>|
|[<span data-ttu-id="bd479-305">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-305">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-306">ReadItem</span></span>|
|[<span data-ttu-id="bd479-307">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-307">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-308">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-309">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-309">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="bd479-310">dateTimeModified: data</span><span class="sxs-lookup"><span data-stu-id="bd479-310">dateTimeModified: Date</span></span>

<span data-ttu-id="bd479-311">Obtém a data e a hora em que um item foi alterado pela última vez.</span><span class="sxs-lookup"><span data-stu-id="bd479-311">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="bd479-312">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="bd479-312">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-313">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="bd479-313">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-314">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-314">Type</span></span>

*   <span data-ttu-id="bd479-315">Data</span><span class="sxs-lookup"><span data-stu-id="bd479-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-316">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-316">Requirements</span></span>

|<span data-ttu-id="bd479-317">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-317">Requirement</span></span>|<span data-ttu-id="bd479-318">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-319">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-320">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-320">1.0</span></span>|
|[<span data-ttu-id="bd479-321">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-321">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-322">ReadItem</span></span>|
|[<span data-ttu-id="bd479-323">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-323">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-324">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-325">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-325">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="bd479-326">fim: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-326">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-327">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="bd479-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="bd479-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="bd479-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bd479-330">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="bd479-330">Read mode</span></span>

<span data-ttu-id="bd479-331">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="bd479-331">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="bd479-332">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="bd479-332">Compose mode</span></span>

<span data-ttu-id="bd479-333">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="bd479-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="bd479-334">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="bd479-334">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="bd479-335">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="bd479-335">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

##### <a name="type"></a><span data-ttu-id="bd479-336">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-336">Type</span></span>

*   <span data-ttu-id="bd479-337">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-337">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-338">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-338">Requirements</span></span>

|<span data-ttu-id="bd479-339">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-339">Requirement</span></span>|<span data-ttu-id="bd479-340">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-341">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-342">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-342">1.0</span></span>|
|[<span data-ttu-id="bd479-343">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-344">ReadItem</span></span>|
|[<span data-ttu-id="bd479-345">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="bd479-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-346">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-346">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17fromjavascriptapioutlookofficefromviewoutlook-js-17"></a><span data-ttu-id="bd479-347">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[de](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-347">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-348">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="bd479-348">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="bd479-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="bd479-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-351">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="bd479-351">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bd479-352">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="bd479-352">Read mode</span></span>

<span data-ttu-id="bd479-353">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="bd479-353">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="bd479-354">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="bd479-354">Compose mode</span></span>

<span data-ttu-id="bd479-355">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="bd479-355">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bd479-356">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-356">Type</span></span>

*   <span data-ttu-id="bd479-357">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [de](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-357">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-358">Requirements</span></span>

|<span data-ttu-id="bd479-359">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-359">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="bd479-360">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-361">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-361">1.0</span></span>|<span data-ttu-id="bd479-362">1.7</span><span class="sxs-lookup"><span data-stu-id="bd479-362">1.7</span></span>|
|[<span data-ttu-id="bd479-363">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-364">ReadItem</span></span>|<span data-ttu-id="bd479-365">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bd479-365">ReadWriteItem</span></span>|
|[<span data-ttu-id="bd479-366">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-367">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-367">Read</span></span>|<span data-ttu-id="bd479-368">Escrever</span><span class="sxs-lookup"><span data-stu-id="bd479-368">Compose</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="bd479-369">internetMessageId: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bd479-369">internetMessageId: String</span></span>

<span data-ttu-id="bd479-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="bd479-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-372">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-372">Type</span></span>

*   <span data-ttu-id="bd479-373">String</span><span class="sxs-lookup"><span data-stu-id="bd479-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-374">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-374">Requirements</span></span>

|<span data-ttu-id="bd479-375">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-375">Requirement</span></span>|<span data-ttu-id="bd479-376">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-377">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-378">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-378">1.0</span></span>|
|[<span data-ttu-id="bd479-379">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-380">ReadItem</span></span>|
|[<span data-ttu-id="bd479-381">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-382">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-383">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-383">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="bd479-384">doclass: String</span><span class="sxs-lookup"><span data-stu-id="bd479-384">itemClass: String</span></span>

<span data-ttu-id="bd479-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="bd479-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="bd479-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="bd479-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="bd479-389">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-389">Type</span></span>|<span data-ttu-id="bd479-390">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-390">Description</span></span>|<span data-ttu-id="bd479-391">classe de item</span><span class="sxs-lookup"><span data-stu-id="bd479-391">item class</span></span>|
|---|---|---|
|<span data-ttu-id="bd479-392">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="bd479-392">Appointment items</span></span>|<span data-ttu-id="bd479-393">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="bd479-393">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="bd479-394">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="bd479-394">Message items</span></span>|<span data-ttu-id="bd479-395">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="bd479-395">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="bd479-396">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="bd479-396">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-397">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-397">Type</span></span>

*   <span data-ttu-id="bd479-398">String</span><span class="sxs-lookup"><span data-stu-id="bd479-398">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-399">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-399">Requirements</span></span>

|<span data-ttu-id="bd479-400">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-400">Requirement</span></span>|<span data-ttu-id="bd479-401">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-402">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-403">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-403">1.0</span></span>|
|[<span data-ttu-id="bd479-404">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-405">ReadItem</span></span>|
|[<span data-ttu-id="bd479-406">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-407">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-407">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-408">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-408">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="bd479-409">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="bd479-409">(nullable) itemId: String</span></span>

<span data-ttu-id="bd479-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="bd479-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-412">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="bd479-412">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="bd479-413">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="bd479-413">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="bd479-414">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="bd479-414">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="bd479-415">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="bd479-415">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="bd479-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-418">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-418">Type</span></span>

*   <span data-ttu-id="bd479-419">String</span><span class="sxs-lookup"><span data-stu-id="bd479-419">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-420">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-420">Requirements</span></span>

|<span data-ttu-id="bd479-421">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-421">Requirement</span></span>|<span data-ttu-id="bd479-422">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-423">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-424">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-424">1.0</span></span>|
|[<span data-ttu-id="bd479-425">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-426">ReadItem</span></span>|
|[<span data-ttu-id="bd479-427">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-427">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-428">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-429">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-429">Example</span></span>

<span data-ttu-id="bd479-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="bd479-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

<br>

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-17"></a><span data-ttu-id="bd479-432">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-432">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-433">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="bd479-433">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="bd479-434">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="bd479-434">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-435">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-435">Type</span></span>

*   [<span data-ttu-id="bd479-436">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="bd479-436">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="bd479-437">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-437">Requirements</span></span>

|<span data-ttu-id="bd479-438">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-438">Requirement</span></span>|<span data-ttu-id="bd479-439">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-440">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-441">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-441">1.0</span></span>|
|[<span data-ttu-id="bd479-442">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-443">ReadItem</span></span>|
|[<span data-ttu-id="bd479-444">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="bd479-444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-445">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-445">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-446">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-446">Example</span></span>

```js
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

<br>

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-17"></a><span data-ttu-id="bd479-447">local: cadeia de caracteres | [Local](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-447">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-448">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="bd479-448">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bd479-449">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="bd479-449">Read mode</span></span>

<span data-ttu-id="bd479-450">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="bd479-450">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="bd479-451">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="bd479-451">Compose mode</span></span>

<span data-ttu-id="bd479-452">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="bd479-452">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bd479-453">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-453">Type</span></span>

*   <span data-ttu-id="bd479-454">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-454">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-455">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-455">Requirements</span></span>

|<span data-ttu-id="bd479-456">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-456">Requirement</span></span>|<span data-ttu-id="bd479-457">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-458">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-459">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-459">1.0</span></span>|
|[<span data-ttu-id="bd479-460">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-461">ReadItem</span></span>|
|[<span data-ttu-id="bd479-462">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="bd479-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-463">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-463">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="bd479-464">normalizedSubject: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bd479-464">normalizedSubject: String</span></span>

<span data-ttu-id="bd479-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="bd479-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="bd479-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="bd479-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-469">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-469">Type</span></span>

*   <span data-ttu-id="bd479-470">String</span><span class="sxs-lookup"><span data-stu-id="bd479-470">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-471">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-471">Requirements</span></span>

|<span data-ttu-id="bd479-472">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-472">Requirement</span></span>|<span data-ttu-id="bd479-473">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-474">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-475">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-475">1.0</span></span>|
|[<span data-ttu-id="bd479-476">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-476">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-477">ReadItem</span></span>|
|[<span data-ttu-id="bd479-478">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-478">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-479">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-479">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-480">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-480">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-17"></a><span data-ttu-id="bd479-481">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-481">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-482">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="bd479-482">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-483">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-483">Type</span></span>

*   [<span data-ttu-id="bd479-484">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="bd479-484">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="bd479-485">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-485">Requirements</span></span>

|<span data-ttu-id="bd479-486">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-486">Requirement</span></span>|<span data-ttu-id="bd479-487">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-488">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-489">1.3</span><span class="sxs-lookup"><span data-stu-id="bd479-489">1.3</span></span>|
|[<span data-ttu-id="bd479-490">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-491">ReadItem</span></span>|
|[<span data-ttu-id="bd479-492">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="bd479-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-493">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-493">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-494">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-494">Example</span></span>

```js
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="bd479-495">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.7) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="bd479-495">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-496">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="bd479-496">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="bd479-497">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="bd479-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bd479-498">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="bd479-498">Read mode</span></span>

<span data-ttu-id="bd479-499">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="bd479-499">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="bd479-500">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="bd479-500">Compose mode</span></span>

<span data-ttu-id="bd479-501">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="bd479-501">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bd479-502">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-502">Type</span></span>

*   <span data-ttu-id="bd479-503">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-503">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-504">Requirements</span></span>

|<span data-ttu-id="bd479-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-505">Requirement</span></span>|<span data-ttu-id="bd479-506">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-507">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-508">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-508">1.0</span></span>|
|[<span data-ttu-id="bd479-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-510">ReadItem</span></span>|
|[<span data-ttu-id="bd479-511">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="bd479-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-512">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-512">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17organizerjavascriptapioutlookofficeorganizerviewoutlook-js-17"></a><span data-ttu-id="bd479-513">organizador: [](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bd479-513">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-514">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="bd479-514">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bd479-515">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="bd479-515">Read mode</span></span>

<span data-ttu-id="bd479-516">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="bd479-516">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="bd479-517">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="bd479-517">Compose mode</span></span>

<span data-ttu-id="bd479-518">A `organizer` propriedade retorna um objeto [organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="bd479-518">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="bd479-519">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-519">Type</span></span>

*   <span data-ttu-id="bd479-520">[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bd479-520">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-521">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-521">Requirements</span></span>

|<span data-ttu-id="bd479-522">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-522">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="bd479-523">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-524">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-524">1.0</span></span>|<span data-ttu-id="bd479-525">1.7</span><span class="sxs-lookup"><span data-stu-id="bd479-525">1.7</span></span>|
|[<span data-ttu-id="bd479-526">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-527">ReadItem</span></span>|<span data-ttu-id="bd479-528">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bd479-528">ReadWriteItem</span></span>|
|[<span data-ttu-id="bd479-529">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-530">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-530">Read</span></span>|<span data-ttu-id="bd479-531">Escrever</span><span class="sxs-lookup"><span data-stu-id="bd479-531">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-17"></a><span data-ttu-id="bd479-532">(anulável) recorrência: [recorrência](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-532">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-533">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="bd479-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="bd479-534">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="bd479-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="bd479-535">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="bd479-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="bd479-536">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="bd479-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="bd479-537">A `recurrence` propriedade retorna um objeto de [recorrência](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="bd479-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="bd479-538">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="bd479-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="bd479-539">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="bd479-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="bd479-540">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="bd479-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="bd479-541">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="bd479-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bd479-542">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="bd479-542">Read mode</span></span>

<span data-ttu-id="bd479-543">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="bd479-543">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that represents the appointment recurrence.</span></span> <span data-ttu-id="bd479-544">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="bd479-544">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="bd479-545">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="bd479-545">Compose mode</span></span>

<span data-ttu-id="bd479-546">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="bd479-546">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="bd479-547">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="bd479-547">This is available for appointments.</span></span>

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a><span data-ttu-id="bd479-548">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-548">Type</span></span>

* [<span data-ttu-id="bd479-549">Recorrência</span><span class="sxs-lookup"><span data-stu-id="bd479-549">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)

|<span data-ttu-id="bd479-550">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-550">Requirement</span></span>|<span data-ttu-id="bd479-551">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-551">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-552">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-552">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-553">1.7</span><span class="sxs-lookup"><span data-stu-id="bd479-553">1.7</span></span>|
|[<span data-ttu-id="bd479-554">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-554">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-555">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-555">ReadItem</span></span>|
|[<span data-ttu-id="bd479-556">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="bd479-556">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-557">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-557">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="bd479-558">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.7) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="bd479-558">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-559">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="bd479-559">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="bd479-560">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="bd479-560">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bd479-561">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="bd479-561">Read mode</span></span>

<span data-ttu-id="bd479-562">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="bd479-562">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="bd479-563">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="bd479-563">Compose mode</span></span>

<span data-ttu-id="bd479-564">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="bd479-564">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="bd479-565">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-565">Type</span></span>

*   <span data-ttu-id="bd479-566">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-566">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-567">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-567">Requirements</span></span>

|<span data-ttu-id="bd479-568">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-568">Requirement</span></span>|<span data-ttu-id="bd479-569">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-570">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-571">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-571">1.0</span></span>|
|[<span data-ttu-id="bd479-572">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-572">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-573">ReadItem</span></span>|
|[<span data-ttu-id="bd479-574">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="bd479-574">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-575">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-575">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17"></a><span data-ttu-id="bd479-576">remetente: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-576">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-p128">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="bd479-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="bd479-p129">As propriedades [`from`](#from-emailaddressdetailsfrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="bd479-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-581">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="bd479-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-582">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-582">Type</span></span>

*   [<span data-ttu-id="bd479-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bd479-583">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="bd479-584">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-584">Requirements</span></span>

|<span data-ttu-id="bd479-585">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-585">Requirement</span></span>|<span data-ttu-id="bd479-586">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-587">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-588">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-588">1.0</span></span>|
|[<span data-ttu-id="bd479-589">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-589">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-590">ReadItem</span></span>|
|[<span data-ttu-id="bd479-591">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-591">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-592">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-593">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-593">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="bd479-594">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="bd479-594">(nullable) seriesId: String</span></span>

<span data-ttu-id="bd479-595">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="bd479-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="bd479-596">No Outlook na Web e clientes de desktop, o `seriesId` retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="bd479-596">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="bd479-597">No entanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="bd479-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-598">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="bd479-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="bd479-599">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="bd479-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="bd479-600">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="bd479-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="bd479-601">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="bd479-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="bd479-602">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="bd479-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="bd479-603">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-603">Type</span></span>

* <span data-ttu-id="bd479-604">String</span><span class="sxs-lookup"><span data-stu-id="bd479-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-605">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-605">Requirements</span></span>

|<span data-ttu-id="bd479-606">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-606">Requirement</span></span>|<span data-ttu-id="bd479-607">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-608">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-609">1.7</span><span class="sxs-lookup"><span data-stu-id="bd479-609">1.7</span></span>|
|[<span data-ttu-id="bd479-610">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-611">ReadItem</span></span>|
|[<span data-ttu-id="bd479-612">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-613">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-613">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-614">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-614">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="bd479-615">Início: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-615">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-616">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="bd479-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="bd479-p132">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="bd479-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bd479-619">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="bd479-619">Read mode</span></span>

<span data-ttu-id="bd479-620">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="bd479-620">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="bd479-621">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="bd479-621">Compose mode</span></span>

<span data-ttu-id="bd479-622">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="bd479-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="bd479-623">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="bd479-623">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="bd479-624">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="bd479-624">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

##### <a name="type"></a><span data-ttu-id="bd479-625">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-625">Type</span></span>

*   <span data-ttu-id="bd479-626">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-626">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-627">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-627">Requirements</span></span>

|<span data-ttu-id="bd479-628">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-628">Requirement</span></span>|<span data-ttu-id="bd479-629">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-629">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-630">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-630">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-631">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-631">1.0</span></span>|
|[<span data-ttu-id="bd479-632">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-632">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-633">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-633">ReadItem</span></span>|
|[<span data-ttu-id="bd479-634">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-634">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-635">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-635">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-17"></a><span data-ttu-id="bd479-636">subject: cadeia de caracteres | [Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-636">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-637">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="bd479-637">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="bd479-638">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="bd479-638">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bd479-639">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="bd479-639">Read mode</span></span>

<span data-ttu-id="bd479-p133">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="bd479-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="bd479-642">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="bd479-642">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="bd479-643">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="bd479-643">Compose mode</span></span>

<span data-ttu-id="bd479-644">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="bd479-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="bd479-645">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-645">Type</span></span>

*   <span data-ttu-id="bd479-646">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-646">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-647">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-647">Requirements</span></span>

|<span data-ttu-id="bd479-648">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-648">Requirement</span></span>|<span data-ttu-id="bd479-649">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-650">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-651">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-651">1.0</span></span>|
|[<span data-ttu-id="bd479-652">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-653">ReadItem</span></span>|
|[<span data-ttu-id="bd479-654">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-655">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-655">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="bd479-656">para: Array. <[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bd479-656">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="bd479-657">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="bd479-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="bd479-658">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="bd479-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bd479-659">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="bd479-659">Read mode</span></span>

<span data-ttu-id="bd479-p135">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="bd479-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="bd479-662">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="bd479-662">Compose mode</span></span>

<span data-ttu-id="bd479-663">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="bd479-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bd479-664">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-664">Type</span></span>

*   <span data-ttu-id="bd479-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-666">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-666">Requirements</span></span>

|<span data-ttu-id="bd479-667">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-667">Requirement</span></span>|<span data-ttu-id="bd479-668">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-669">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-670">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-670">1.0</span></span>|
|[<span data-ttu-id="bd479-671">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-672">ReadItem</span></span>|
|[<span data-ttu-id="bd479-673">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-674">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-674">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="bd479-675">Métodos</span><span class="sxs-lookup"><span data-stu-id="bd479-675">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="bd479-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bd479-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="bd479-677">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="bd479-677">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="bd479-678">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="bd479-678">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="bd479-679">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="bd479-679">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-680">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-680">Parameters</span></span>
|<span data-ttu-id="bd479-681">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-681">Name</span></span>|<span data-ttu-id="bd479-682">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-682">Type</span></span>|<span data-ttu-id="bd479-683">Atributos</span><span class="sxs-lookup"><span data-stu-id="bd479-683">Attributes</span></span>|<span data-ttu-id="bd479-684">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-684">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="bd479-685">String</span><span class="sxs-lookup"><span data-stu-id="bd479-685">String</span></span>||<span data-ttu-id="bd479-p136">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bd479-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="bd479-688">String</span><span class="sxs-lookup"><span data-stu-id="bd479-688">String</span></span>||<span data-ttu-id="bd479-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bd479-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="bd479-691">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-691">Object</span></span>|<span data-ttu-id="bd479-692">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-692">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-693">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="bd479-693">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="bd479-694">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-694">Object</span></span>|<span data-ttu-id="bd479-695">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-695">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-696">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-696">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="bd479-697">Booliano</span><span class="sxs-lookup"><span data-stu-id="bd479-697">Boolean</span></span>|<span data-ttu-id="bd479-698">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-698">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-699">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="bd479-699">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="bd479-700">function</span><span class="sxs-lookup"><span data-stu-id="bd479-700">function</span></span>|<span data-ttu-id="bd479-701">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-701">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-702">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd479-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bd479-703">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bd479-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="bd479-704">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="bd479-704">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bd479-705">Erros</span><span class="sxs-lookup"><span data-stu-id="bd479-705">Errors</span></span>

|<span data-ttu-id="bd479-706">Código de erro</span><span class="sxs-lookup"><span data-stu-id="bd479-706">Error code</span></span>|<span data-ttu-id="bd479-707">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-707">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="bd479-708">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="bd479-708">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="bd479-709">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="bd479-709">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="bd479-710">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="bd479-710">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-711">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-711">Requirements</span></span>

|<span data-ttu-id="bd479-712">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-712">Requirement</span></span>|<span data-ttu-id="bd479-713">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-713">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-714">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-714">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-715">1.1</span><span class="sxs-lookup"><span data-stu-id="bd479-715">1.1</span></span>|
|[<span data-ttu-id="bd479-716">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-716">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-717">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bd479-717">ReadWriteItem</span></span>|
|[<span data-ttu-id="bd479-718">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-718">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-719">Escrever</span><span class="sxs-lookup"><span data-stu-id="bd479-719">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="bd479-720">Exemplos</span><span class="sxs-lookup"><span data-stu-id="bd479-720">Examples</span></span>

```js
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

<span data-ttu-id="bd479-721">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="bd479-721">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
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

<br>

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="bd479-722">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bd479-722">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="bd479-723">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="bd479-723">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="bd479-724">Atualmente, os tipos de eventos `Office.EventType.AppointmentTimeChanged`suportados `Office.EventType.RecipientsChanged`são, e`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="bd479-724">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-725">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-725">Parameters</span></span>

| <span data-ttu-id="bd479-726">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-726">Name</span></span> | <span data-ttu-id="bd479-727">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-727">Type</span></span> | <span data-ttu-id="bd479-728">Atributos</span><span class="sxs-lookup"><span data-stu-id="bd479-728">Attributes</span></span> | <span data-ttu-id="bd479-729">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-729">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="bd479-730">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="bd479-730">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="bd479-731">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="bd479-731">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="bd479-732">Função</span><span class="sxs-lookup"><span data-stu-id="bd479-732">Function</span></span> || <span data-ttu-id="bd479-p138">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="bd479-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="bd479-736">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-736">Object</span></span> | <span data-ttu-id="bd479-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-737">&lt;optional&gt;</span></span> | <span data-ttu-id="bd479-738">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="bd479-738">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="bd479-739">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-739">Object</span></span> | <span data-ttu-id="bd479-740">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-740">&lt;optional&gt;</span></span> | <span data-ttu-id="bd479-741">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-741">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="bd479-742">function</span><span class="sxs-lookup"><span data-stu-id="bd479-742">function</span></span>| <span data-ttu-id="bd479-743">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-743">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-744">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd479-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-745">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-745">Requirements</span></span>

|<span data-ttu-id="bd479-746">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-746">Requirement</span></span>| <span data-ttu-id="bd479-747">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-748">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd479-749">1.7</span><span class="sxs-lookup"><span data-stu-id="bd479-749">1.7</span></span> |
|[<span data-ttu-id="bd479-750">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd479-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-751">ReadItem</span></span> |
|[<span data-ttu-id="bd479-752">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="bd479-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd479-753">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-753">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="bd479-754">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-754">Example</span></span>

```js
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="bd479-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bd479-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="bd479-756">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="bd479-756">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="bd479-p139">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="bd479-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="bd479-760">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="bd479-760">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="bd479-761">Se o suplemento do Office estiver em execução no Outlook na Web, o `addItemAttachmentAsync` método poderá anexar itens a itens diferentes do item que você está editando; no entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="bd479-761">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-762">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-762">Parameters</span></span>

|<span data-ttu-id="bd479-763">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-763">Name</span></span>|<span data-ttu-id="bd479-764">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-764">Type</span></span>|<span data-ttu-id="bd479-765">Atributos</span><span class="sxs-lookup"><span data-stu-id="bd479-765">Attributes</span></span>|<span data-ttu-id="bd479-766">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-766">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="bd479-767">String</span><span class="sxs-lookup"><span data-stu-id="bd479-767">String</span></span>||<span data-ttu-id="bd479-p140">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bd479-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="bd479-770">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bd479-770">String</span></span>||<span data-ttu-id="bd479-771">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="bd479-771">The subject of the item to be attached.</span></span> <span data-ttu-id="bd479-772">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bd479-772">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="bd479-773">Object</span><span class="sxs-lookup"><span data-stu-id="bd479-773">Object</span></span>|<span data-ttu-id="bd479-774">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-774">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-775">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="bd479-775">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="bd479-776">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-776">Object</span></span>|<span data-ttu-id="bd479-777">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-777">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-778">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-778">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="bd479-779">function</span><span class="sxs-lookup"><span data-stu-id="bd479-779">function</span></span>|<span data-ttu-id="bd479-780">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-780">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-781">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd479-781">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bd479-782">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bd479-782">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="bd479-783">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="bd479-783">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bd479-784">Erros</span><span class="sxs-lookup"><span data-stu-id="bd479-784">Errors</span></span>

|<span data-ttu-id="bd479-785">Código de erro</span><span class="sxs-lookup"><span data-stu-id="bd479-785">Error code</span></span>|<span data-ttu-id="bd479-786">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-786">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="bd479-787">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="bd479-787">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-788">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-788">Requirements</span></span>

|<span data-ttu-id="bd479-789">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-789">Requirement</span></span>|<span data-ttu-id="bd479-790">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-791">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-792">1.1</span><span class="sxs-lookup"><span data-stu-id="bd479-792">1.1</span></span>|
|[<span data-ttu-id="bd479-793">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-794">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bd479-794">ReadWriteItem</span></span>|
|[<span data-ttu-id="bd479-795">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-796">Escrever</span><span class="sxs-lookup"><span data-stu-id="bd479-796">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-797">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-797">Example</span></span>

<span data-ttu-id="bd479-798">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="bd479-798">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
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

<br>

---
---

#### <a name="close"></a><span data-ttu-id="bd479-799">close()</span><span class="sxs-lookup"><span data-stu-id="bd479-799">close()</span></span>

<span data-ttu-id="bd479-800">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="bd479-800">Closes the current item that is being composed.</span></span>

<span data-ttu-id="bd479-p142">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="bd479-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-803">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="bd479-803">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="bd479-804">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="bd479-804">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-805">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-805">Requirements</span></span>

|<span data-ttu-id="bd479-806">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-806">Requirement</span></span>|<span data-ttu-id="bd479-807">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-808">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-809">1.3</span><span class="sxs-lookup"><span data-stu-id="bd479-809">1.3</span></span>|
|[<span data-ttu-id="bd479-810">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-811">Restrito</span><span class="sxs-lookup"><span data-stu-id="bd479-811">Restricted</span></span>|
|[<span data-ttu-id="bd479-812">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-813">Escrever</span><span class="sxs-lookup"><span data-stu-id="bd479-813">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="bd479-814">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="bd479-814">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="bd479-815">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="bd479-815">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-816">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="bd479-816">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bd479-817">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="bd479-817">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="bd479-818">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="bd479-818">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="bd479-819">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="bd479-819">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="bd479-820">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="bd479-820">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="bd479-821">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="bd479-821">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-822">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-822">Parameters</span></span>

|<span data-ttu-id="bd479-823">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-823">Name</span></span>|<span data-ttu-id="bd479-824">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-824">Type</span></span>|<span data-ttu-id="bd479-825">Atributos</span><span class="sxs-lookup"><span data-stu-id="bd479-825">Attributes</span></span>|<span data-ttu-id="bd479-826">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-826">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="bd479-827">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="bd479-827">String &#124; Object</span></span>||<span data-ttu-id="bd479-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="bd479-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="bd479-830">**OU**</span><span class="sxs-lookup"><span data-stu-id="bd479-830">**OR**</span></span><br/><span data-ttu-id="bd479-p145">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="bd479-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="bd479-833">String</span><span class="sxs-lookup"><span data-stu-id="bd479-833">String</span></span>|<span data-ttu-id="bd479-834">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-834">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="bd479-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="bd479-837">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-837">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="bd479-838">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-838">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-839">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="bd479-839">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="bd479-840">String</span><span class="sxs-lookup"><span data-stu-id="bd479-840">String</span></span>||<span data-ttu-id="bd479-p147">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="bd479-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="bd479-843">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bd479-843">String</span></span>||<span data-ttu-id="bd479-844">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="bd479-844">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="bd479-845">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bd479-845">String</span></span>||<span data-ttu-id="bd479-p148">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="bd479-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="bd479-848">Booliano</span><span class="sxs-lookup"><span data-stu-id="bd479-848">Boolean</span></span>||<span data-ttu-id="bd479-p149">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="bd479-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="bd479-851">String</span><span class="sxs-lookup"><span data-stu-id="bd479-851">String</span></span>||<span data-ttu-id="bd479-p150">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bd479-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="bd479-855">function</span><span class="sxs-lookup"><span data-stu-id="bd479-855">function</span></span>|<span data-ttu-id="bd479-856">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-856">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-857">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd479-857">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-858">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-858">Requirements</span></span>

|<span data-ttu-id="bd479-859">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-859">Requirement</span></span>|<span data-ttu-id="bd479-860">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-861">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-862">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-862">1.0</span></span>|
|[<span data-ttu-id="bd479-863">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-863">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-864">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-864">ReadItem</span></span>|
|[<span data-ttu-id="bd479-865">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-865">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-866">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-866">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="bd479-867">Exemplos</span><span class="sxs-lookup"><span data-stu-id="bd479-867">Examples</span></span>

<span data-ttu-id="bd479-868">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="bd479-868">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="bd479-869">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="bd479-869">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="bd479-870">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="bd479-870">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="bd479-871">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="bd479-871">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="bd479-872">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="bd479-872">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="bd479-873">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-873">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="bd479-874">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="bd479-874">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="bd479-875">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="bd479-875">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-876">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="bd479-876">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bd479-877">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="bd479-877">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="bd479-878">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="bd479-878">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="bd479-879">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="bd479-879">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="bd479-880">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="bd479-880">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="bd479-881">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="bd479-881">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-882">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-882">Parameters</span></span>

|<span data-ttu-id="bd479-883">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-883">Name</span></span>|<span data-ttu-id="bd479-884">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-884">Type</span></span>|<span data-ttu-id="bd479-885">Atributos</span><span class="sxs-lookup"><span data-stu-id="bd479-885">Attributes</span></span>|<span data-ttu-id="bd479-886">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-886">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="bd479-887">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="bd479-887">String &#124; Object</span></span>||<span data-ttu-id="bd479-p152">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="bd479-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="bd479-890">**OU**</span><span class="sxs-lookup"><span data-stu-id="bd479-890">**OR**</span></span><br/><span data-ttu-id="bd479-p153">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="bd479-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="bd479-893">String</span><span class="sxs-lookup"><span data-stu-id="bd479-893">String</span></span>|<span data-ttu-id="bd479-894">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-894">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="bd479-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="bd479-897">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-897">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="bd479-898">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-898">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-899">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="bd479-899">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="bd479-900">String</span><span class="sxs-lookup"><span data-stu-id="bd479-900">String</span></span>||<span data-ttu-id="bd479-p155">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="bd479-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="bd479-903">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bd479-903">String</span></span>||<span data-ttu-id="bd479-904">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="bd479-904">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="bd479-905">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bd479-905">String</span></span>||<span data-ttu-id="bd479-p156">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="bd479-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="bd479-908">Booliano</span><span class="sxs-lookup"><span data-stu-id="bd479-908">Boolean</span></span>||<span data-ttu-id="bd479-p157">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="bd479-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="bd479-911">String</span><span class="sxs-lookup"><span data-stu-id="bd479-911">String</span></span>||<span data-ttu-id="bd479-p158">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bd479-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="bd479-915">function</span><span class="sxs-lookup"><span data-stu-id="bd479-915">function</span></span>|<span data-ttu-id="bd479-916">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-916">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-917">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd479-917">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-918">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-918">Requirements</span></span>

|<span data-ttu-id="bd479-919">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-919">Requirement</span></span>|<span data-ttu-id="bd479-920">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-920">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-921">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-921">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-922">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-922">1.0</span></span>|
|[<span data-ttu-id="bd479-923">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-923">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-924">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-924">ReadItem</span></span>|
|[<span data-ttu-id="bd479-925">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-925">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-926">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-926">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="bd479-927">Exemplos</span><span class="sxs-lookup"><span data-stu-id="bd479-927">Examples</span></span>

<span data-ttu-id="bd479-928">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="bd479-928">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="bd479-929">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="bd479-929">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="bd479-930">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="bd479-930">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="bd479-931">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="bd479-931">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="bd479-932">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="bd479-932">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="bd479-933">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-933">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="bd479-934">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="bd479-934">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="bd479-935">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="bd479-935">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-936">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="bd479-936">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-937">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-937">Requirements</span></span>

|<span data-ttu-id="bd479-938">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-938">Requirement</span></span>|<span data-ttu-id="bd479-939">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-940">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-941">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-941">1.0</span></span>|
|[<span data-ttu-id="bd479-942">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-942">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-943">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-943">ReadItem</span></span>|
|[<span data-ttu-id="bd479-944">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-944">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-945">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-945">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bd479-946">Retorna:</span><span class="sxs-lookup"><span data-stu-id="bd479-946">Returns:</span></span>

<span data-ttu-id="bd479-947">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-947">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="bd479-948">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-948">Example</span></span>

<span data-ttu-id="bd479-949">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="bd479-949">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="bd479-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="bd479-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="bd479-951">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="bd479-951">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-952">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="bd479-952">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-953">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-953">Parameters</span></span>

|<span data-ttu-id="bd479-954">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-954">Name</span></span>|<span data-ttu-id="bd479-955">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-955">Type</span></span>|<span data-ttu-id="bd479-956">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-956">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="bd479-957">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="bd479-957">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.7)|<span data-ttu-id="bd479-958">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="bd479-958">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-959">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-959">Requirements</span></span>

|<span data-ttu-id="bd479-960">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-960">Requirement</span></span>|<span data-ttu-id="bd479-961">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-962">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-963">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-963">1.0</span></span>|
|[<span data-ttu-id="bd479-964">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-964">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-965">Restrito</span><span class="sxs-lookup"><span data-stu-id="bd479-965">Restricted</span></span>|
|[<span data-ttu-id="bd479-966">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-966">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-967">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bd479-968">Retorna:</span><span class="sxs-lookup"><span data-stu-id="bd479-968">Returns:</span></span>

<span data-ttu-id="bd479-969">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="bd479-969">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="bd479-970">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="bd479-970">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="bd479-971">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="bd479-971">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="bd479-972">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="bd479-972">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="bd479-973">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="bd479-973">Value of `entityType`</span></span>|<span data-ttu-id="bd479-974">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="bd479-974">Type of objects in returned array</span></span>|<span data-ttu-id="bd479-975">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="bd479-975">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="bd479-976">String</span><span class="sxs-lookup"><span data-stu-id="bd479-976">String</span></span>|<span data-ttu-id="bd479-977">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="bd479-977">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="bd479-978">Contato</span><span class="sxs-lookup"><span data-stu-id="bd479-978">Contact</span></span>|<span data-ttu-id="bd479-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bd479-979">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="bd479-980">String</span><span class="sxs-lookup"><span data-stu-id="bd479-980">String</span></span>|<span data-ttu-id="bd479-981">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bd479-981">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="bd479-982">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="bd479-982">MeetingSuggestion</span></span>|<span data-ttu-id="bd479-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bd479-983">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="bd479-984">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="bd479-984">PhoneNumber</span></span>|<span data-ttu-id="bd479-985">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="bd479-985">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="bd479-986">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="bd479-986">TaskSuggestion</span></span>|<span data-ttu-id="bd479-987">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bd479-987">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="bd479-988">String</span><span class="sxs-lookup"><span data-stu-id="bd479-988">String</span></span>|<span data-ttu-id="bd479-989">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="bd479-989">**Restricted**</span></span>|

<span data-ttu-id="bd479-990">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="bd479-990">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

##### <a name="example"></a><span data-ttu-id="bd479-991">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-991">Example</span></span>

<span data-ttu-id="bd479-992">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="bd479-992">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="bd479-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="bd479-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="bd479-994">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="bd479-994">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-995">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="bd479-995">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bd479-996">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="bd479-996">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-997">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-997">Parameters</span></span>

|<span data-ttu-id="bd479-998">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-998">Name</span></span>|<span data-ttu-id="bd479-999">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-999">Type</span></span>|<span data-ttu-id="bd479-1000">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-1000">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="bd479-1001">String</span><span class="sxs-lookup"><span data-stu-id="bd479-1001">String</span></span>|<span data-ttu-id="bd479-1002">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="bd479-1002">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-1003">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-1003">Requirements</span></span>

|<span data-ttu-id="bd479-1004">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-1004">Requirement</span></span>|<span data-ttu-id="bd479-1005">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-1006">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-1007">1.0</span></span>|
|[<span data-ttu-id="bd479-1008">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-1008">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-1009">ReadItem</span></span>|
|[<span data-ttu-id="bd479-1010">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-1010">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-1011">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bd479-1012">Retorna:</span><span class="sxs-lookup"><span data-stu-id="bd479-1012">Returns:</span></span>

<span data-ttu-id="bd479-p160">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="bd479-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="bd479-1015">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="bd479-1015">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="bd479-1016">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="bd479-1016">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="bd479-1017">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="bd479-1017">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-1018">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="bd479-1018">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bd479-p161">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="bd479-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="bd479-1022">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="bd479-1022">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="bd479-1023">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="bd479-1023">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="bd479-p162">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="bd479-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-1027">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-1027">Requirements</span></span>

|<span data-ttu-id="bd479-1028">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-1028">Requirement</span></span>|<span data-ttu-id="bd479-1029">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-1029">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-1030">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-1030">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-1031">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-1031">1.0</span></span>|
|[<span data-ttu-id="bd479-1032">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-1032">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-1033">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-1033">ReadItem</span></span>|
|[<span data-ttu-id="bd479-1034">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-1034">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-1035">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-1035">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bd479-1036">Retorna:</span><span class="sxs-lookup"><span data-stu-id="bd479-1036">Returns:</span></span>

<span data-ttu-id="bd479-p163">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="bd479-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="bd479-1039">Tipo: objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-1039">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="bd479-1040">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-1040">Example</span></span>

<span data-ttu-id="bd479-1041">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="bd479-1041">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="bd479-1042">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="bd479-1042">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="bd479-1043">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="bd479-1043">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-1044">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="bd479-1044">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bd479-1045">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="bd479-1045">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="bd479-p164">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="bd479-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-1048">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-1048">Parameters</span></span>

|<span data-ttu-id="bd479-1049">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-1049">Name</span></span>|<span data-ttu-id="bd479-1050">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-1050">Type</span></span>|<span data-ttu-id="bd479-1051">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-1051">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="bd479-1052">String</span><span class="sxs-lookup"><span data-stu-id="bd479-1052">String</span></span>|<span data-ttu-id="bd479-1053">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="bd479-1053">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-1054">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-1054">Requirements</span></span>

|<span data-ttu-id="bd479-1055">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-1055">Requirement</span></span>|<span data-ttu-id="bd479-1056">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-1056">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-1057">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-1057">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-1058">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-1058">1.0</span></span>|
|[<span data-ttu-id="bd479-1059">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-1059">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-1060">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-1060">ReadItem</span></span>|
|[<span data-ttu-id="bd479-1061">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-1061">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-1062">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-1062">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bd479-1063">Retorna:</span><span class="sxs-lookup"><span data-stu-id="bd479-1063">Returns:</span></span>

<span data-ttu-id="bd479-1064">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="bd479-1064">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="bd479-1065">Tipo: cadeia de caracteres de matriz. < ></span><span class="sxs-lookup"><span data-stu-id="bd479-1065">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="bd479-1066">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-1066">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="bd479-1067">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="bd479-1067">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="bd479-1068">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="bd479-1068">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="bd479-p165">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="bd479-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-1071">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-1071">Parameters</span></span>

|<span data-ttu-id="bd479-1072">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-1072">Name</span></span>|<span data-ttu-id="bd479-1073">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-1073">Type</span></span>|<span data-ttu-id="bd479-1074">Atributos</span><span class="sxs-lookup"><span data-stu-id="bd479-1074">Attributes</span></span>|<span data-ttu-id="bd479-1075">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-1075">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="bd479-1076">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="bd479-1076">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="bd479-p166">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="bd479-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="bd479-1080">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-1080">Object</span></span>|<span data-ttu-id="bd479-1081">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1081">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-1082">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="bd479-1082">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="bd479-1083">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-1083">Object</span></span>|<span data-ttu-id="bd479-1084">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1084">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-1085">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-1085">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="bd479-1086">function</span><span class="sxs-lookup"><span data-stu-id="bd479-1086">function</span></span>||<span data-ttu-id="bd479-1087">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd479-1087">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bd479-1088">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="bd479-1088">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="bd479-1089">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="bd479-1089">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-1090">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-1090">Requirements</span></span>

|<span data-ttu-id="bd479-1091">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-1091">Requirement</span></span>|<span data-ttu-id="bd479-1092">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-1092">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-1093">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-1093">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-1094">1.2</span><span class="sxs-lookup"><span data-stu-id="bd479-1094">1.2</span></span>|
|[<span data-ttu-id="bd479-1095">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-1095">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-1096">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-1096">ReadItem</span></span>|
|[<span data-ttu-id="bd479-1097">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-1097">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-1098">Escrever</span><span class="sxs-lookup"><span data-stu-id="bd479-1098">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="bd479-1099">Retorna:</span><span class="sxs-lookup"><span data-stu-id="bd479-1099">Returns:</span></span>

<span data-ttu-id="bd479-1100">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="bd479-1100">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="bd479-1101">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="bd479-1101">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="bd479-1102">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-1102">Example</span></span>

```js
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

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="bd479-1103">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="bd479-1103">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="bd479-1104">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="bd479-1104">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="bd479-1105">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="bd479-1105">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-1106">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="bd479-1106">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-1107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-1107">Requirements</span></span>

|<span data-ttu-id="bd479-1108">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-1108">Requirement</span></span>|<span data-ttu-id="bd479-1109">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-1109">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-1110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-1110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-1111">1.6</span><span class="sxs-lookup"><span data-stu-id="bd479-1111">1.6</span></span>|
|[<span data-ttu-id="bd479-1112">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-1112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-1113">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-1113">ReadItem</span></span>|
|[<span data-ttu-id="bd479-1114">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-1114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-1115">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-1115">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bd479-1116">Retorna:</span><span class="sxs-lookup"><span data-stu-id="bd479-1116">Returns:</span></span>

<span data-ttu-id="bd479-1117">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="bd479-1117">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="bd479-1118">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-1118">Example</span></span>

<span data-ttu-id="bd479-1119">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="bd479-1119">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="bd479-1120">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="bd479-1120">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="bd479-p169">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="bd479-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-1123">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="bd479-1123">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bd479-p170">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="bd479-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="bd479-1127">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="bd479-1127">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="bd479-1128">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="bd479-1128">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="bd479-p171">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="bd479-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd479-1132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-1132">Requirements</span></span>

|<span data-ttu-id="bd479-1133">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-1133">Requirement</span></span>|<span data-ttu-id="bd479-1134">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-1135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-1136">1.6</span><span class="sxs-lookup"><span data-stu-id="bd479-1136">1.6</span></span>|
|[<span data-ttu-id="bd479-1137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-1137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-1138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-1138">ReadItem</span></span>|
|[<span data-ttu-id="bd479-1139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-1139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-1140">Read</span><span class="sxs-lookup"><span data-stu-id="bd479-1140">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bd479-1141">Retorna:</span><span class="sxs-lookup"><span data-stu-id="bd479-1141">Returns:</span></span>

<span data-ttu-id="bd479-p172">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="bd479-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="bd479-1144">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-1144">Example</span></span>

<span data-ttu-id="bd479-1145">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="bd479-1145">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="bd479-1146">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="bd479-1146">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="bd479-1147">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="bd479-1147">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="bd479-p173">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="bd479-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-1151">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-1151">Parameters</span></span>

|<span data-ttu-id="bd479-1152">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-1152">Name</span></span>|<span data-ttu-id="bd479-1153">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-1153">Type</span></span>|<span data-ttu-id="bd479-1154">Atributos</span><span class="sxs-lookup"><span data-stu-id="bd479-1154">Attributes</span></span>|<span data-ttu-id="bd479-1155">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-1155">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="bd479-1156">function</span><span class="sxs-lookup"><span data-stu-id="bd479-1156">function</span></span>||<span data-ttu-id="bd479-1157">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd479-1157">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bd479-1158">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bd479-1158">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="bd479-1159">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="bd479-1159">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="bd479-1160">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-1160">Object</span></span>|<span data-ttu-id="bd479-1161">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1161">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-1162">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-1162">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="bd479-1163">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-1163">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-1164">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-1164">Requirements</span></span>

|<span data-ttu-id="bd479-1165">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-1165">Requirement</span></span>|<span data-ttu-id="bd479-1166">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-1166">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-1167">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-1167">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-1168">1.0</span><span class="sxs-lookup"><span data-stu-id="bd479-1168">1.0</span></span>|
|[<span data-ttu-id="bd479-1169">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-1169">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-1170">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-1170">ReadItem</span></span>|
|[<span data-ttu-id="bd479-1171">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="bd479-1171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-1172">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-1172">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-1173">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-1173">Example</span></span>

<span data-ttu-id="bd479-p176">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="bd479-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
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

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="bd479-1177">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bd479-1177">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="bd479-1178">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="bd479-1178">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="bd479-1179">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="bd479-1179">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="bd479-1180">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="bd479-1180">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="bd479-1181">No Outlook na Web e dispositivos móveis, o identificador de anexo é válido somente dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="bd479-1181">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="bd479-1182">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="bd479-1182">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-1183">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-1183">Parameters</span></span>

|<span data-ttu-id="bd479-1184">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-1184">Name</span></span>|<span data-ttu-id="bd479-1185">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-1185">Type</span></span>|<span data-ttu-id="bd479-1186">Atributos</span><span class="sxs-lookup"><span data-stu-id="bd479-1186">Attributes</span></span>|<span data-ttu-id="bd479-1187">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-1187">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="bd479-1188">String</span><span class="sxs-lookup"><span data-stu-id="bd479-1188">String</span></span>||<span data-ttu-id="bd479-1189">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="bd479-1189">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="bd479-1190">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-1190">Object</span></span>|<span data-ttu-id="bd479-1191">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1191">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-1192">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="bd479-1192">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="bd479-1193">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-1193">Object</span></span>|<span data-ttu-id="bd479-1194">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1194">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-1195">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-1195">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="bd479-1196">function</span><span class="sxs-lookup"><span data-stu-id="bd479-1196">function</span></span>|<span data-ttu-id="bd479-1197">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1197">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-1198">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd479-1198">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bd479-1199">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="bd479-1199">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bd479-1200">Erros</span><span class="sxs-lookup"><span data-stu-id="bd479-1200">Errors</span></span>

|<span data-ttu-id="bd479-1201">Código de erro</span><span class="sxs-lookup"><span data-stu-id="bd479-1201">Error code</span></span>|<span data-ttu-id="bd479-1202">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-1202">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="bd479-1203">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="bd479-1203">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-1204">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-1204">Requirements</span></span>

|<span data-ttu-id="bd479-1205">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-1205">Requirement</span></span>|<span data-ttu-id="bd479-1206">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-1206">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-1207">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-1207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-1208">1.1</span><span class="sxs-lookup"><span data-stu-id="bd479-1208">1.1</span></span>|
|[<span data-ttu-id="bd479-1209">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-1209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-1210">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bd479-1210">ReadWriteItem</span></span>|
|[<span data-ttu-id="bd479-1211">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-1211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-1212">Escrever</span><span class="sxs-lookup"><span data-stu-id="bd479-1212">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-1213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-1213">Example</span></span>

<span data-ttu-id="bd479-1214">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="bd479-1214">The following code removes an attachment with an identifier of '0'.</span></span>

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

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="bd479-1215">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bd479-1215">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="bd479-1216">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="bd479-1216">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="bd479-1217">Atualmente, os tipos de eventos `Office.EventType.AppointmentTimeChanged`suportados `Office.EventType.RecipientsChanged`são, e`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="bd479-1217">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-1218">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-1218">Parameters</span></span>

| <span data-ttu-id="bd479-1219">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-1219">Name</span></span> | <span data-ttu-id="bd479-1220">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-1220">Type</span></span> | <span data-ttu-id="bd479-1221">Atributos</span><span class="sxs-lookup"><span data-stu-id="bd479-1221">Attributes</span></span> | <span data-ttu-id="bd479-1222">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-1222">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="bd479-1223">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="bd479-1223">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="bd479-1224">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="bd479-1224">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="bd479-1225">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-1225">Object</span></span> | <span data-ttu-id="bd479-1226">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1226">&lt;optional&gt;</span></span> | <span data-ttu-id="bd479-1227">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="bd479-1227">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="bd479-1228">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-1228">Object</span></span> | <span data-ttu-id="bd479-1229">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="bd479-1230">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-1230">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="bd479-1231">function</span><span class="sxs-lookup"><span data-stu-id="bd479-1231">function</span></span>| <span data-ttu-id="bd479-1232">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1232">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-1233">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd479-1233">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-1234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-1234">Requirements</span></span>

|<span data-ttu-id="bd479-1235">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-1235">Requirement</span></span>| <span data-ttu-id="bd479-1236">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-1236">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-1237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-1237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd479-1238">1.7</span><span class="sxs-lookup"><span data-stu-id="bd479-1238">1.7</span></span> |
|[<span data-ttu-id="bd479-1239">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-1239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd479-1240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd479-1240">ReadItem</span></span> |
|[<span data-ttu-id="bd479-1241">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="bd479-1241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd479-1242">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bd479-1242">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="bd479-1243">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-1243">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};
```

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="bd479-1244">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="bd479-1244">saveAsync([options], callback)</span></span>

<span data-ttu-id="bd479-1245">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="bd479-1245">Asynchronously saves an item.</span></span>

<span data-ttu-id="bd479-1246">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-1246">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="bd479-1247">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="bd479-1247">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="bd479-1248">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="bd479-1248">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-1249">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="bd479-1249">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="bd479-1250">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="bd479-1250">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="bd479-p180">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="bd479-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="bd479-1254">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="bd479-1254">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="bd479-1255">O Outlook no Mac não dá suporte à gravação de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="bd479-1255">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="bd479-1256">O `saveAsync` método falha quando chamado de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="bd479-1256">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="bd479-1257">Consulte [não é possível salvar uma reunião como rascunho no Outlook para Mac usando a API do Office js](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="bd479-1257">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="bd479-1258">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="bd479-1258">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-1259">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-1259">Parameters</span></span>

|<span data-ttu-id="bd479-1260">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-1260">Name</span></span>|<span data-ttu-id="bd479-1261">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-1261">Type</span></span>|<span data-ttu-id="bd479-1262">Atributos</span><span class="sxs-lookup"><span data-stu-id="bd479-1262">Attributes</span></span>|<span data-ttu-id="bd479-1263">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-1263">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="bd479-1264">Object</span><span class="sxs-lookup"><span data-stu-id="bd479-1264">Object</span></span>|<span data-ttu-id="bd479-1265">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-1266">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="bd479-1266">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="bd479-1267">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-1267">Object</span></span>|<span data-ttu-id="bd479-1268">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-1269">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-1269">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="bd479-1270">function</span><span class="sxs-lookup"><span data-stu-id="bd479-1270">function</span></span>||<span data-ttu-id="bd479-1271">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd479-1271">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bd479-1272">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bd479-1272">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-1273">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-1273">Requirements</span></span>

|<span data-ttu-id="bd479-1274">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-1274">Requirement</span></span>|<span data-ttu-id="bd479-1275">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-1275">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-1276">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-1276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-1277">1.3</span><span class="sxs-lookup"><span data-stu-id="bd479-1277">1.3</span></span>|
|[<span data-ttu-id="bd479-1278">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-1278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-1279">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bd479-1279">ReadWriteItem</span></span>|
|[<span data-ttu-id="bd479-1280">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-1280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-1281">Escrever</span><span class="sxs-lookup"><span data-stu-id="bd479-1281">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="bd479-1282">Exemplos</span><span class="sxs-lookup"><span data-stu-id="bd479-1282">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="bd479-p182">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="bd479-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="bd479-1285">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="bd479-1285">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="bd479-1286">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="bd479-1286">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="bd479-p183">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="bd479-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd479-1290">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="bd479-1290">Parameters</span></span>

|<span data-ttu-id="bd479-1291">Nome</span><span class="sxs-lookup"><span data-stu-id="bd479-1291">Name</span></span>|<span data-ttu-id="bd479-1292">Tipo</span><span class="sxs-lookup"><span data-stu-id="bd479-1292">Type</span></span>|<span data-ttu-id="bd479-1293">Atributos</span><span class="sxs-lookup"><span data-stu-id="bd479-1293">Attributes</span></span>|<span data-ttu-id="bd479-1294">Descrição</span><span class="sxs-lookup"><span data-stu-id="bd479-1294">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="bd479-1295">String</span><span class="sxs-lookup"><span data-stu-id="bd479-1295">String</span></span>||<span data-ttu-id="bd479-p184">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="bd479-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="bd479-1299">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-1299">Object</span></span>|<span data-ttu-id="bd479-1300">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1300">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-1301">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="bd479-1301">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="bd479-1302">Objeto</span><span class="sxs-lookup"><span data-stu-id="bd479-1302">Object</span></span>|<span data-ttu-id="bd479-1303">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1303">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-1304">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="bd479-1304">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="bd479-1305">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="bd479-1305">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="bd479-1306">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="bd479-1306">&lt;optional&gt;</span></span>|<span data-ttu-id="bd479-1307">Se `text`, o estilo atual é aplicado no Outlook na Web e clientes da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="bd479-1307">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="bd479-1308">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="bd479-1308">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="bd479-1309">Se `html` e o campo oferecer suporte a HTML (o assunto não), o estilo atual será aplicado no Outlook na Web e o estilo padrão será aplicado nos clientes da área de trabalho do Outlook.</span><span class="sxs-lookup"><span data-stu-id="bd479-1309">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="bd479-1310">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="bd479-1310">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="bd479-1311">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="bd479-1311">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="bd479-1312">function</span><span class="sxs-lookup"><span data-stu-id="bd479-1312">function</span></span>||<span data-ttu-id="bd479-1313">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd479-1313">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd479-1314">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bd479-1314">Requirements</span></span>

|<span data-ttu-id="bd479-1315">Requisito</span><span class="sxs-lookup"><span data-stu-id="bd479-1315">Requirement</span></span>|<span data-ttu-id="bd479-1316">Valor</span><span class="sxs-lookup"><span data-stu-id="bd479-1316">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd479-1317">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bd479-1317">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="bd479-1318">1.2</span><span class="sxs-lookup"><span data-stu-id="bd479-1318">1.2</span></span>|
|[<span data-ttu-id="bd479-1319">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bd479-1319">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="bd479-1320">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bd479-1320">ReadWriteItem</span></span>|
|[<span data-ttu-id="bd479-1321">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bd479-1321">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="bd479-1322">Escrever</span><span class="sxs-lookup"><span data-stu-id="bd479-1322">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bd479-1323">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bd479-1323">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
