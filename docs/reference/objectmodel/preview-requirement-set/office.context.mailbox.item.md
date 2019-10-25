---
title: Office. Context. Mailbox. Item-visualização do conjunto de requisitos
description: ''
ms.date: 10/23/2019
localization_priority: Normal
ms.openlocfilehash: 7a72e6fbbec6dbf9cee07d85237baef93ca7ecd8
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682659"
---
# <a name="item"></a><span data-ttu-id="28b08-102">item</span><span class="sxs-lookup"><span data-stu-id="28b08-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="28b08-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="28b08-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="28b08-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="28b08-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-106">Requirements</span></span>

|<span data-ttu-id="28b08-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-107">Requirement</span></span>|<span data-ttu-id="28b08-108">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-110">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-110">1.0</span></span>|
|[<span data-ttu-id="28b08-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="28b08-112">Restricted</span></span>|
|[<span data-ttu-id="28b08-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="28b08-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="28b08-115">Members and methods</span></span>

| <span data-ttu-id="28b08-116">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-116">Member</span></span> | <span data-ttu-id="28b08-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="28b08-118">attachments</span><span class="sxs-lookup"><span data-stu-id="28b08-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="28b08-119">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-119">Member</span></span> |
| [<span data-ttu-id="28b08-120">bcc</span><span class="sxs-lookup"><span data-stu-id="28b08-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="28b08-121">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-121">Member</span></span> |
| [<span data-ttu-id="28b08-122">body</span><span class="sxs-lookup"><span data-stu-id="28b08-122">body</span></span>](#body-body) | <span data-ttu-id="28b08-123">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-123">Member</span></span> |
| [<span data-ttu-id="28b08-124">categories</span><span class="sxs-lookup"><span data-stu-id="28b08-124">categories</span></span>](#categories-categories) | <span data-ttu-id="28b08-125">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-125">Member</span></span> |
| [<span data-ttu-id="28b08-126">cc</span><span class="sxs-lookup"><span data-stu-id="28b08-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="28b08-127">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-127">Member</span></span> |
| [<span data-ttu-id="28b08-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="28b08-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="28b08-129">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-129">Member</span></span> |
| [<span data-ttu-id="28b08-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="28b08-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="28b08-131">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-131">Member</span></span> |
| [<span data-ttu-id="28b08-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="28b08-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="28b08-133">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-133">Member</span></span> |
| [<span data-ttu-id="28b08-134">end</span><span class="sxs-lookup"><span data-stu-id="28b08-134">end</span></span>](#end-datetime) | <span data-ttu-id="28b08-135">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-135">Member</span></span> |
| [<span data-ttu-id="28b08-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="28b08-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="28b08-137">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-137">Member</span></span> |
| [<span data-ttu-id="28b08-138">from</span><span class="sxs-lookup"><span data-stu-id="28b08-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="28b08-139">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-139">Member</span></span> |
| [<span data-ttu-id="28b08-140">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="28b08-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="28b08-141">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-141">Member</span></span> |
| [<span data-ttu-id="28b08-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="28b08-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="28b08-143">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-143">Member</span></span> |
| [<span data-ttu-id="28b08-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="28b08-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="28b08-145">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-145">Member</span></span> |
| [<span data-ttu-id="28b08-146">itemId</span><span class="sxs-lookup"><span data-stu-id="28b08-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="28b08-147">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-147">Member</span></span> |
| [<span data-ttu-id="28b08-148">itemType</span><span class="sxs-lookup"><span data-stu-id="28b08-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="28b08-149">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-149">Member</span></span> |
| [<span data-ttu-id="28b08-150">location</span><span class="sxs-lookup"><span data-stu-id="28b08-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="28b08-151">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-151">Member</span></span> |
| [<span data-ttu-id="28b08-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="28b08-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="28b08-153">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-153">Member</span></span> |
| [<span data-ttu-id="28b08-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="28b08-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="28b08-155">Member</span><span class="sxs-lookup"><span data-stu-id="28b08-155">Member</span></span> |
| [<span data-ttu-id="28b08-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="28b08-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="28b08-157">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-157">Member</span></span> |
| [<span data-ttu-id="28b08-158">organizer</span><span class="sxs-lookup"><span data-stu-id="28b08-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="28b08-159">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-159">Member</span></span> |
| [<span data-ttu-id="28b08-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="28b08-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="28b08-161">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-161">Member</span></span> |
| [<span data-ttu-id="28b08-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="28b08-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="28b08-163">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-163">Member</span></span> |
| [<span data-ttu-id="28b08-164">sender</span><span class="sxs-lookup"><span data-stu-id="28b08-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="28b08-165">Member</span><span class="sxs-lookup"><span data-stu-id="28b08-165">Member</span></span> |
| [<span data-ttu-id="28b08-166">seriesid</span><span class="sxs-lookup"><span data-stu-id="28b08-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="28b08-167">Member</span><span class="sxs-lookup"><span data-stu-id="28b08-167">Member</span></span> |
| [<span data-ttu-id="28b08-168">start</span><span class="sxs-lookup"><span data-stu-id="28b08-168">start</span></span>](#start-datetime) | <span data-ttu-id="28b08-169">Member</span><span class="sxs-lookup"><span data-stu-id="28b08-169">Member</span></span> |
| [<span data-ttu-id="28b08-170">subject</span><span class="sxs-lookup"><span data-stu-id="28b08-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="28b08-171">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-171">Member</span></span> |
| [<span data-ttu-id="28b08-172">to</span><span class="sxs-lookup"><span data-stu-id="28b08-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="28b08-173">Membro</span><span class="sxs-lookup"><span data-stu-id="28b08-173">Member</span></span> |
| [<span data-ttu-id="28b08-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="28b08-175">Método</span><span class="sxs-lookup"><span data-stu-id="28b08-175">Method</span></span> |
| [<span data-ttu-id="28b08-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="28b08-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="28b08-177">Método</span><span class="sxs-lookup"><span data-stu-id="28b08-177">Method</span></span> |
| [<span data-ttu-id="28b08-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="28b08-179">Método</span><span class="sxs-lookup"><span data-stu-id="28b08-179">Method</span></span> |
| [<span data-ttu-id="28b08-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="28b08-181">Método</span><span class="sxs-lookup"><span data-stu-id="28b08-181">Method</span></span> |
| [<span data-ttu-id="28b08-182">close</span><span class="sxs-lookup"><span data-stu-id="28b08-182">close</span></span>](#close) | <span data-ttu-id="28b08-183">Método</span><span class="sxs-lookup"><span data-stu-id="28b08-183">Method</span></span> |
| [<span data-ttu-id="28b08-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="28b08-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="28b08-185">Método</span><span class="sxs-lookup"><span data-stu-id="28b08-185">Method</span></span> |
| [<span data-ttu-id="28b08-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="28b08-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="28b08-187">Método</span><span class="sxs-lookup"><span data-stu-id="28b08-187">Method</span></span> |
| [<span data-ttu-id="28b08-188">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-188">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="28b08-189">Método</span><span class="sxs-lookup"><span data-stu-id="28b08-189">Method</span></span> |
| [<span data-ttu-id="28b08-190">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-190">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="28b08-191">Método</span><span class="sxs-lookup"><span data-stu-id="28b08-191">Method</span></span> |
| [<span data-ttu-id="28b08-192">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-192">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="28b08-193">Método</span><span class="sxs-lookup"><span data-stu-id="28b08-193">Method</span></span> |
| [<span data-ttu-id="28b08-194">getEntities</span><span class="sxs-lookup"><span data-stu-id="28b08-194">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="28b08-195">Method</span><span class="sxs-lookup"><span data-stu-id="28b08-195">Method</span></span> |
| [<span data-ttu-id="28b08-196">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="28b08-196">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="28b08-197">Method</span><span class="sxs-lookup"><span data-stu-id="28b08-197">Method</span></span> |
| [<span data-ttu-id="28b08-198">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="28b08-198">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="28b08-199">Method</span><span class="sxs-lookup"><span data-stu-id="28b08-199">Method</span></span> |
| [<span data-ttu-id="28b08-200">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-200">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="28b08-201">Method</span><span class="sxs-lookup"><span data-stu-id="28b08-201">Method</span></span> |
| [<span data-ttu-id="28b08-202">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-202">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="28b08-203">Method</span><span class="sxs-lookup"><span data-stu-id="28b08-203">Method</span></span> |
| [<span data-ttu-id="28b08-204">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="28b08-204">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="28b08-205">Method</span><span class="sxs-lookup"><span data-stu-id="28b08-205">Method</span></span> |
| [<span data-ttu-id="28b08-206">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="28b08-206">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="28b08-207">Method</span><span class="sxs-lookup"><span data-stu-id="28b08-207">Method</span></span> |
| [<span data-ttu-id="28b08-208">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-208">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="28b08-209">Method</span><span class="sxs-lookup"><span data-stu-id="28b08-209">Method</span></span> |
| [<span data-ttu-id="28b08-210">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="28b08-210">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="28b08-211">Method</span><span class="sxs-lookup"><span data-stu-id="28b08-211">Method</span></span> |
| [<span data-ttu-id="28b08-212">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="28b08-212">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="28b08-213">Method</span><span class="sxs-lookup"><span data-stu-id="28b08-213">Method</span></span> |
| [<span data-ttu-id="28b08-214">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-214">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="28b08-215">Method</span><span class="sxs-lookup"><span data-stu-id="28b08-215">Method</span></span> |
| [<span data-ttu-id="28b08-216">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-216">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="28b08-217">Method</span><span class="sxs-lookup"><span data-stu-id="28b08-217">Method</span></span> |
| [<span data-ttu-id="28b08-218">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-218">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="28b08-219">Método</span><span class="sxs-lookup"><span data-stu-id="28b08-219">Method</span></span> |
| [<span data-ttu-id="28b08-220">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-220">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="28b08-221">Método</span><span class="sxs-lookup"><span data-stu-id="28b08-221">Method</span></span> |
| [<span data-ttu-id="28b08-222">saveAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-222">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="28b08-223">Method</span><span class="sxs-lookup"><span data-stu-id="28b08-223">Method</span></span> |
| [<span data-ttu-id="28b08-224">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="28b08-224">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="28b08-225">Método</span><span class="sxs-lookup"><span data-stu-id="28b08-225">Method</span></span> |

### <a name="example"></a><span data-ttu-id="28b08-226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-226">Example</span></span>

<span data-ttu-id="28b08-227">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="28b08-227">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="28b08-228">Members</span><span class="sxs-lookup"><span data-stu-id="28b08-228">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="28b08-229">anexos: Matriz.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="28b08-229">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="28b08-230">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="28b08-230">Gets the item's attachments as an array.</span></span> <span data-ttu-id="28b08-231">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="28b08-231">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-232">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="28b08-232">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="28b08-233">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="28b08-233">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-234">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-234">Type</span></span>

*   <span data-ttu-id="28b08-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="28b08-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-236">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-236">Requirements</span></span>

|<span data-ttu-id="28b08-237">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-237">Requirement</span></span>|<span data-ttu-id="28b08-238">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-239">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-240">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-240">1.0</span></span>|
|[<span data-ttu-id="28b08-241">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-242">ReadItem</span></span>|
|[<span data-ttu-id="28b08-243">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-244">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-244">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-245">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-245">Example</span></span>

<span data-ttu-id="28b08-246">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="28b08-246">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="28b08-247">cco :[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="28b08-247">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="28b08-248">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="28b08-248">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="28b08-249">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="28b08-249">Compose mode only.</span></span>

<span data-ttu-id="28b08-250">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="28b08-251">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="28b08-251">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="28b08-252">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-252">Get 500 members maximum.</span></span>
- <span data-ttu-id="28b08-253">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="28b08-253">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-254">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-254">Type</span></span>

*   [<span data-ttu-id="28b08-255">Destinatários</span><span class="sxs-lookup"><span data-stu-id="28b08-255">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="28b08-256">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-256">Requirements</span></span>

|<span data-ttu-id="28b08-257">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-257">Requirement</span></span>|<span data-ttu-id="28b08-258">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-259">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-260">1.1</span><span class="sxs-lookup"><span data-stu-id="28b08-260">1.1</span></span>|
|[<span data-ttu-id="28b08-261">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-262">ReadItem</span></span>|
|[<span data-ttu-id="28b08-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-264">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-264">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-265">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-265">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="28b08-266">corpo: [Corpo](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="28b08-266">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="28b08-267">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="28b08-267">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-268">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-268">Type</span></span>

*   [<span data-ttu-id="28b08-269">Body</span><span class="sxs-lookup"><span data-stu-id="28b08-269">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="28b08-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-270">Requirements</span></span>

|<span data-ttu-id="28b08-271">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-271">Requirement</span></span>|<span data-ttu-id="28b08-272">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-273">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-274">1.1</span><span class="sxs-lookup"><span data-stu-id="28b08-274">1.1</span></span>|
|[<span data-ttu-id="28b08-275">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-275">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-276">ReadItem</span></span>|
|[<span data-ttu-id="28b08-277">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-277">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-278">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-279">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-279">Example</span></span>

<span data-ttu-id="28b08-280">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="28b08-280">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="28b08-281">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-281">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="28b08-282">Categorias: [categorias](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="28b08-282">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="28b08-283">Obtém um objeto que fornece métodos para gerenciar as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="28b08-283">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-284">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="28b08-284">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-285">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-285">Type</span></span>

*   [<span data-ttu-id="28b08-286">Categories</span><span class="sxs-lookup"><span data-stu-id="28b08-286">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="28b08-287">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-287">Requirements</span></span>

|<span data-ttu-id="28b08-288">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-288">Requirement</span></span>|<span data-ttu-id="28b08-289">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-290">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-291">Visualização</span><span class="sxs-lookup"><span data-stu-id="28b08-291">Preview</span></span>|
|[<span data-ttu-id="28b08-292">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-293">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-293">ReadItem</span></span>|
|[<span data-ttu-id="28b08-294">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-295">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-295">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-296">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-296">Example</span></span>

<span data-ttu-id="28b08-297">Este exemplo obtém as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="28b08-297">This example gets the item's categories.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="28b08-298">cc : Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="28b08-298">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="28b08-299">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="28b08-299">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="28b08-300">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="28b08-300">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="28b08-301">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="28b08-301">Read mode</span></span>

<span data-ttu-id="28b08-302">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="28b08-302">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="28b08-303">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-303">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="28b08-304">No entanto, no Windows e no Mac, é possível obter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-304">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="28b08-305">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="28b08-305">Compose mode</span></span>

<span data-ttu-id="28b08-306">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="28b08-306">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="28b08-307">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-307">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="28b08-308">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="28b08-308">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="28b08-309">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-309">Get 500 members maximum.</span></span>
- <span data-ttu-id="28b08-310">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="28b08-310">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="28b08-311">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-311">Type</span></span>

*   <span data-ttu-id="28b08-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="28b08-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-313">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-313">Requirements</span></span>

|<span data-ttu-id="28b08-314">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-314">Requirement</span></span>|<span data-ttu-id="28b08-315">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-316">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-317">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-317">1.0</span></span>|
|[<span data-ttu-id="28b08-318">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-319">ReadItem</span></span>|
|[<span data-ttu-id="28b08-320">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-321">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-321">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="28b08-322">(anulável) conversationId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-322">(nullable) conversationId: String</span></span>

<span data-ttu-id="28b08-323">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="28b08-323">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="28b08-p109">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="28b08-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="28b08-p110">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="28b08-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-328">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-328">Type</span></span>

*   <span data-ttu-id="28b08-329">String</span><span class="sxs-lookup"><span data-stu-id="28b08-329">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-330">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-330">Requirements</span></span>

|<span data-ttu-id="28b08-331">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-331">Requirement</span></span>|<span data-ttu-id="28b08-332">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-333">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-334">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-334">1.0</span></span>|
|[<span data-ttu-id="28b08-335">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-336">ReadItem</span></span>|
|[<span data-ttu-id="28b08-337">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-338">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-338">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-339">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-339">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="28b08-340">dateTimeCreated: Data</span><span class="sxs-lookup"><span data-stu-id="28b08-340">dateTimeCreated: Date</span></span>

<span data-ttu-id="28b08-p111">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="28b08-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-343">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-343">Type</span></span>

*   <span data-ttu-id="28b08-344">Data</span><span class="sxs-lookup"><span data-stu-id="28b08-344">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-345">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-345">Requirements</span></span>

|<span data-ttu-id="28b08-346">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-346">Requirement</span></span>|<span data-ttu-id="28b08-347">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-348">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-349">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-349">1.0</span></span>|
|[<span data-ttu-id="28b08-350">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-350">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-351">ReadItem</span></span>|
|[<span data-ttu-id="28b08-352">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-352">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-353">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-353">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-354">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-354">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="28b08-355">dateTimeModified: Data</span><span class="sxs-lookup"><span data-stu-id="28b08-355">dateTimeModified: Date</span></span>

<span data-ttu-id="28b08-p112">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="28b08-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-358">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="28b08-358">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-359">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-359">Type</span></span>

*   <span data-ttu-id="28b08-360">Data</span><span class="sxs-lookup"><span data-stu-id="28b08-360">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-361">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-361">Requirements</span></span>

|<span data-ttu-id="28b08-362">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-362">Requirement</span></span>|<span data-ttu-id="28b08-363">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-363">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-364">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-364">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-365">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-365">1.0</span></span>|
|[<span data-ttu-id="28b08-366">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-366">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-367">ReadItem</span></span>|
|[<span data-ttu-id="28b08-368">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-368">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-369">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-369">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-370">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-370">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="28b08-371">fim: Data|[Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="28b08-371">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="28b08-372">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="28b08-372">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="28b08-p113">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="28b08-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="28b08-375">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="28b08-375">Read mode</span></span>

<span data-ttu-id="28b08-376">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="28b08-376">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="28b08-377">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="28b08-377">Compose mode</span></span>

<span data-ttu-id="28b08-378">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="28b08-378">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="28b08-379">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="28b08-379">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="28b08-380">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="28b08-380">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="28b08-381">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-381">Type</span></span>

*   <span data-ttu-id="28b08-382">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="28b08-382">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-383">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-383">Requirements</span></span>

|<span data-ttu-id="28b08-384">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-384">Requirement</span></span>|<span data-ttu-id="28b08-385">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-386">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-387">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-387">1.0</span></span>|
|[<span data-ttu-id="28b08-388">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-389">ReadItem</span></span>|
|[<span data-ttu-id="28b08-390">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-391">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-391">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="28b08-392">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="28b08-392">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="28b08-393">Obtém ou define os locais de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-393">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="28b08-394">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="28b08-394">Read mode</span></span>

<span data-ttu-id="28b08-395">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que permite que você obtenha o conjunto de locais (cada um representado por um objeto [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associado ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-395">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="28b08-396">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="28b08-396">Compose mode</span></span>

<span data-ttu-id="28b08-397">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que fornece métodos para obter, remover ou adicionar locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-397">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-398">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-398">Type</span></span>

*   [<span data-ttu-id="28b08-399">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="28b08-399">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="28b08-400">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-400">Requirements</span></span>

|<span data-ttu-id="28b08-401">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-401">Requirement</span></span>|<span data-ttu-id="28b08-402">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-403">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-403">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-404">Visualização</span><span class="sxs-lookup"><span data-stu-id="28b08-404">Preview</span></span>|
|[<span data-ttu-id="28b08-405">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-405">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-406">ReadItem</span></span>|
|[<span data-ttu-id="28b08-407">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-407">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-408">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-408">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-409">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-409">Example</span></span>

<span data-ttu-id="28b08-410">O exemplo a seguir obtém os locais atuais associados ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-410">The following example gets the current locations associated with the appointment.</span></span>

```js
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="28b08-411">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="28b08-411">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="28b08-412">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="28b08-412">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="28b08-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="28b08-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-415">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="28b08-415">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="28b08-416">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="28b08-416">Read mode</span></span>

<span data-ttu-id="28b08-417">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="28b08-417">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="28b08-418">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="28b08-418">Compose mode</span></span>

<span data-ttu-id="28b08-419">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="28b08-419">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="28b08-420">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-420">Type</span></span>

*   <span data-ttu-id="28b08-421">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="28b08-421">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-422">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-422">Requirements</span></span>

|<span data-ttu-id="28b08-423">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-423">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="28b08-424">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-425">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-425">1.0</span></span>|<span data-ttu-id="28b08-426">1.7</span><span class="sxs-lookup"><span data-stu-id="28b08-426">1.7</span></span>|
|[<span data-ttu-id="28b08-427">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-427">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-428">ReadItem</span></span>|<span data-ttu-id="28b08-429">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="28b08-429">ReadWriteItem</span></span>|
|[<span data-ttu-id="28b08-430">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-431">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-431">Read</span></span>|<span data-ttu-id="28b08-432">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-432">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="28b08-433">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="28b08-433">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="28b08-434">Obtém ou define cabeçalhos de Internet personalizados em uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="28b08-434">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="28b08-435">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="28b08-435">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-436">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-436">Type</span></span>

*   [<span data-ttu-id="28b08-437">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="28b08-437">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="28b08-438">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-438">Requirements</span></span>

|<span data-ttu-id="28b08-439">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-439">Requirement</span></span>|<span data-ttu-id="28b08-440">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-440">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-441">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-442">Visualização</span><span class="sxs-lookup"><span data-stu-id="28b08-442">Preview</span></span>|
|[<span data-ttu-id="28b08-443">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-444">ReadItem</span></span>|
|[<span data-ttu-id="28b08-445">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-446">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-446">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-447">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-447">Example</span></span>

```js
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="28b08-448">internetMessageId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-448">internetMessageId: String</span></span>

<span data-ttu-id="28b08-p116">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="28b08-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-451">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-451">Type</span></span>

*   <span data-ttu-id="28b08-452">String</span><span class="sxs-lookup"><span data-stu-id="28b08-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-453">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-453">Requirements</span></span>

|<span data-ttu-id="28b08-454">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-454">Requirement</span></span>|<span data-ttu-id="28b08-455">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-456">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-457">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-457">1.0</span></span>|
|[<span data-ttu-id="28b08-458">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-458">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-459">ReadItem</span></span>|
|[<span data-ttu-id="28b08-460">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-460">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-461">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-462">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-462">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="28b08-463">itemClass: Cadeia de caracteres </span><span class="sxs-lookup"><span data-stu-id="28b08-463">itemClass: String</span></span>

<span data-ttu-id="28b08-p117">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="28b08-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="28b08-p118">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="28b08-468">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-468">Type</span></span>|<span data-ttu-id="28b08-469">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-469">Description</span></span>|<span data-ttu-id="28b08-470">classe de item</span><span class="sxs-lookup"><span data-stu-id="28b08-470">item class</span></span>|
|---|---|---|
|<span data-ttu-id="28b08-471">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="28b08-471">Appointment items</span></span>|<span data-ttu-id="28b08-472">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="28b08-472">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="28b08-473">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="28b08-473">Message items</span></span>|<span data-ttu-id="28b08-474">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="28b08-474">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="28b08-475">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="28b08-475">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-476">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-476">Type</span></span>

*   <span data-ttu-id="28b08-477">String</span><span class="sxs-lookup"><span data-stu-id="28b08-477">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-478">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-478">Requirements</span></span>

|<span data-ttu-id="28b08-479">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-479">Requirement</span></span>|<span data-ttu-id="28b08-480">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-481">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-482">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-482">1.0</span></span>|
|[<span data-ttu-id="28b08-483">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-484">ReadItem</span></span>|
|[<span data-ttu-id="28b08-485">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-486">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-486">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-487">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-487">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="28b08-488">(anulável) itemId: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-488">(nullable) itemId: String</span></span>

<span data-ttu-id="28b08-p119">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="28b08-p119">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-491">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="28b08-491">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="28b08-492">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="28b08-492">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="28b08-493">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="28b08-493">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="28b08-494">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="28b08-494">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="28b08-p121">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-497">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-497">Type</span></span>

*   <span data-ttu-id="28b08-498">String</span><span class="sxs-lookup"><span data-stu-id="28b08-498">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-499">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-499">Requirements</span></span>

|<span data-ttu-id="28b08-500">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-500">Requirement</span></span>|<span data-ttu-id="28b08-501">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-502">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-503">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-503">1.0</span></span>|
|[<span data-ttu-id="28b08-504">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-504">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-505">ReadItem</span></span>|
|[<span data-ttu-id="28b08-506">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-506">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-507">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-508">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-508">Example</span></span>

<span data-ttu-id="28b08-p122">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="28b08-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="28b08-511">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="28b08-511">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="28b08-512">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="28b08-512">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="28b08-513">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-513">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-514">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-514">Type</span></span>

*   [<span data-ttu-id="28b08-515">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="28b08-515">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="28b08-516">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-516">Requirements</span></span>

|<span data-ttu-id="28b08-517">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-517">Requirement</span></span>|<span data-ttu-id="28b08-518">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-519">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-520">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-520">1.0</span></span>|
|[<span data-ttu-id="28b08-521">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-521">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-522">ReadItem</span></span>|
|[<span data-ttu-id="28b08-523">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-523">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-524">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-524">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-525">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-525">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="28b08-526">Local: Cadeia de caracteres[Local](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="28b08-526">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="28b08-527">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-527">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="28b08-528">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="28b08-528">Read mode</span></span>

<span data-ttu-id="28b08-529">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-529">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="28b08-530">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="28b08-530">Compose mode</span></span>

<span data-ttu-id="28b08-531">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-531">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="28b08-532">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-532">Type</span></span>

*   <span data-ttu-id="28b08-533">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="28b08-533">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-534">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-534">Requirements</span></span>

|<span data-ttu-id="28b08-535">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-535">Requirement</span></span>|<span data-ttu-id="28b08-536">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-537">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-538">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-538">1.0</span></span>|
|[<span data-ttu-id="28b08-539">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-539">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-540">ReadItem</span></span>|
|[<span data-ttu-id="28b08-541">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-541">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-542">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-542">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="28b08-543">normalizedSubject: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-543">normalizedSubject: String</span></span>

<span data-ttu-id="28b08-p123">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="28b08-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="28b08-p124">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="28b08-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-548">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-548">Type</span></span>

*   <span data-ttu-id="28b08-549">String</span><span class="sxs-lookup"><span data-stu-id="28b08-549">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-550">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-550">Requirements</span></span>

|<span data-ttu-id="28b08-551">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-551">Requirement</span></span>|<span data-ttu-id="28b08-552">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-552">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-553">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-553">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-554">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-554">1.0</span></span>|
|[<span data-ttu-id="28b08-555">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-555">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-556">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-556">ReadItem</span></span>|
|[<span data-ttu-id="28b08-557">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-557">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-558">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-558">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-559">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-559">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="28b08-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="28b08-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="28b08-561">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="28b08-561">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-562">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-562">Type</span></span>

*   [<span data-ttu-id="28b08-563">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="28b08-563">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="28b08-564">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-564">Requirements</span></span>

|<span data-ttu-id="28b08-565">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-565">Requirement</span></span>|<span data-ttu-id="28b08-566">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-567">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-568">1.3</span><span class="sxs-lookup"><span data-stu-id="28b08-568">1.3</span></span>|
|[<span data-ttu-id="28b08-569">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-570">ReadItem</span></span>|
|[<span data-ttu-id="28b08-571">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-572">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-572">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-573">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-573">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="28b08-574">optionalAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="28b08-574">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="28b08-575">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="28b08-575">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="28b08-576">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="28b08-576">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="28b08-577">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="28b08-577">Read mode</span></span>

<span data-ttu-id="28b08-578">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="28b08-578">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="28b08-579">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-579">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="28b08-580">No entanto, no Windows e no Mac, é possível obter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-580">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="28b08-581">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="28b08-581">Compose mode</span></span>

<span data-ttu-id="28b08-582">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="28b08-582">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="28b08-583">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-583">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="28b08-584">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="28b08-584">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="28b08-585">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-585">Get 500 members maximum.</span></span>
- <span data-ttu-id="28b08-586">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="28b08-586">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="28b08-587">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-587">Type</span></span>

*   <span data-ttu-id="28b08-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="28b08-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-589">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-589">Requirements</span></span>

|<span data-ttu-id="28b08-590">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-590">Requirement</span></span>|<span data-ttu-id="28b08-591">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-592">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-593">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-593">1.0</span></span>|
|[<span data-ttu-id="28b08-594">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-595">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-595">ReadItem</span></span>|
|[<span data-ttu-id="28b08-596">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-597">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-597">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="28b08-598">organizador: [](/javascript/api/outlook/office.emailaddressdetails)|[organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="28b08-598">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="28b08-599">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="28b08-599">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="28b08-600">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="28b08-600">Read mode</span></span>

<span data-ttu-id="28b08-601">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="28b08-601">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="28b08-602">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="28b08-602">Compose mode</span></span>

<span data-ttu-id="28b08-603">A `organizer` propriedade retorna um objeto [organizador](/javascript/api/outlook/office.organizer) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="28b08-603">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="28b08-604">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-604">Type</span></span>

*   <span data-ttu-id="28b08-605">[](/javascript/api/outlook/office.emailaddressdetails) | [Organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="28b08-605">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-606">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-606">Requirements</span></span>

|<span data-ttu-id="28b08-607">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-607">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="28b08-608">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-609">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-609">1.0</span></span>|<span data-ttu-id="28b08-610">1.7</span><span class="sxs-lookup"><span data-stu-id="28b08-610">1.7</span></span>|
|[<span data-ttu-id="28b08-611">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-611">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-612">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-612">ReadItem</span></span>|<span data-ttu-id="28b08-613">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="28b08-613">ReadWriteItem</span></span>|
|[<span data-ttu-id="28b08-614">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-615">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-615">Read</span></span>|<span data-ttu-id="28b08-616">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-616">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="28b08-617">(anulável) recorrência: [recorrência](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="28b08-617">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="28b08-618">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-618">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="28b08-619">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="28b08-619">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="28b08-620">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-620">Read and compose modes for appointment items.</span></span> <span data-ttu-id="28b08-621">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="28b08-621">Read mode for meeting request items.</span></span>

<span data-ttu-id="28b08-622">A `recurrence` propriedade retorna um objeto de [recorrência](/javascript/api/outlook/office.recurrence) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="28b08-622">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="28b08-623">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="28b08-623">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="28b08-624">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="28b08-624">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="28b08-625">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="28b08-625">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="28b08-626">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="28b08-626">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="28b08-627">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="28b08-627">Read mode</span></span>

<span data-ttu-id="28b08-628">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-628">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="28b08-629">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="28b08-629">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="28b08-630">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="28b08-630">Compose mode</span></span>

<span data-ttu-id="28b08-631">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-631">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="28b08-632">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="28b08-632">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="28b08-633">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-633">Type</span></span>

* [<span data-ttu-id="28b08-634">Recorrência</span><span class="sxs-lookup"><span data-stu-id="28b08-634">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="28b08-635">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-635">Requirement</span></span>|<span data-ttu-id="28b08-636">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-637">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-638">1.7</span><span class="sxs-lookup"><span data-stu-id="28b08-638">1.7</span></span>|
|[<span data-ttu-id="28b08-639">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-640">ReadItem</span></span>|
|[<span data-ttu-id="28b08-641">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-642">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-642">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="28b08-643">requiredAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="28b08-643">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="28b08-644">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="28b08-644">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="28b08-645">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="28b08-645">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="28b08-646">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="28b08-646">Read mode</span></span>

<span data-ttu-id="28b08-647">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="28b08-647">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="28b08-648">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-648">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="28b08-649">No entanto, no Windows e no Mac, é possível obter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-649">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="28b08-650">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="28b08-650">Compose mode</span></span>

<span data-ttu-id="28b08-651">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="28b08-651">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="28b08-652">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-652">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="28b08-653">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="28b08-653">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="28b08-654">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-654">Get 500 members maximum.</span></span>
- <span data-ttu-id="28b08-655">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="28b08-655">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="28b08-656">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-656">Type</span></span>

*   <span data-ttu-id="28b08-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="28b08-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-658">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-658">Requirements</span></span>

|<span data-ttu-id="28b08-659">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-659">Requirement</span></span>|<span data-ttu-id="28b08-660">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-661">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-662">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-662">1.0</span></span>|
|[<span data-ttu-id="28b08-663">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-664">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-664">ReadItem</span></span>|
|[<span data-ttu-id="28b08-665">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-666">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-666">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="28b08-667">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="28b08-667">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="28b08-p135">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="28b08-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="28b08-p136">As propriedades [`from`](#from-emailaddressdetailsfrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="28b08-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-672">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="28b08-672">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-673">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-673">Type</span></span>

*   [<span data-ttu-id="28b08-674">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="28b08-674">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="28b08-675">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-675">Requirements</span></span>

|<span data-ttu-id="28b08-676">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-676">Requirement</span></span>|<span data-ttu-id="28b08-677">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-677">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-678">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-678">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-679">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-679">1.0</span></span>|
|[<span data-ttu-id="28b08-680">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-680">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-681">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-681">ReadItem</span></span>|
|[<span data-ttu-id="28b08-682">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-682">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-683">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-683">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-684">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-684">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="28b08-685">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="28b08-685">(nullable) seriesId: String</span></span>

<span data-ttu-id="28b08-686">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="28b08-686">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="28b08-687">No Outlook na Web e clientes de desktop, o `seriesId` retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="28b08-687">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="28b08-688">No entanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="28b08-688">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-689">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="28b08-689">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="28b08-690">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="28b08-690">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="28b08-691">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="28b08-691">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="28b08-692">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="28b08-692">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="28b08-693">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="28b08-693">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="28b08-694">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-694">Type</span></span>

* <span data-ttu-id="28b08-695">String</span><span class="sxs-lookup"><span data-stu-id="28b08-695">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-696">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-696">Requirements</span></span>

|<span data-ttu-id="28b08-697">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-697">Requirement</span></span>|<span data-ttu-id="28b08-698">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-699">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-700">1.7</span><span class="sxs-lookup"><span data-stu-id="28b08-700">1.7</span></span>|
|[<span data-ttu-id="28b08-701">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-701">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-702">ReadItem</span></span>|
|[<span data-ttu-id="28b08-703">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-703">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-704">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-704">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-705">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-705">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="28b08-706">início: Data|[Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="28b08-706">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="28b08-707">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="28b08-707">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="28b08-p139">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="28b08-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="28b08-710">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="28b08-710">Read mode</span></span>

<span data-ttu-id="28b08-711">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="28b08-711">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="28b08-712">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="28b08-712">Compose mode</span></span>

<span data-ttu-id="28b08-713">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="28b08-713">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="28b08-714">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="28b08-714">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="28b08-715">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="28b08-715">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="28b08-716">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-716">Type</span></span>

*   <span data-ttu-id="28b08-717">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="28b08-717">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-718">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-718">Requirements</span></span>

|<span data-ttu-id="28b08-719">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-719">Requirement</span></span>|<span data-ttu-id="28b08-720">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-720">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-721">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-721">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-722">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-722">1.0</span></span>|
|[<span data-ttu-id="28b08-723">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-723">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-724">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-724">ReadItem</span></span>|
|[<span data-ttu-id="28b08-725">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-725">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-726">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-726">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="28b08-727">Assunto: Cadeia de caracteres|[Assunto](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="28b08-727">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="28b08-728">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="28b08-728">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="28b08-729">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="28b08-729">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="28b08-730">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="28b08-730">Read mode</span></span>

<span data-ttu-id="28b08-p140">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="28b08-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="28b08-733">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="28b08-733">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="28b08-734">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="28b08-734">Compose mode</span></span>
<span data-ttu-id="28b08-735">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="28b08-735">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="28b08-736">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-736">Type</span></span>

*   <span data-ttu-id="28b08-737">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="28b08-737">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-738">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-738">Requirements</span></span>

|<span data-ttu-id="28b08-739">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-739">Requirement</span></span>|<span data-ttu-id="28b08-740">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-741">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-741">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-742">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-742">1.0</span></span>|
|[<span data-ttu-id="28b08-743">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-743">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-744">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-744">ReadItem</span></span>|
|[<span data-ttu-id="28b08-745">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-745">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-746">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-746">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="28b08-747">para: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="28b08-747">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="28b08-748">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="28b08-748">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="28b08-749">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="28b08-749">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="28b08-750">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="28b08-750">Read mode</span></span>

<span data-ttu-id="28b08-751">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="28b08-751">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="28b08-752">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-752">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="28b08-753">No entanto, no Windows e no Mac, é possível obter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-753">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="28b08-754">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="28b08-754">Compose mode</span></span>

<span data-ttu-id="28b08-755">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="28b08-755">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="28b08-756">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-756">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="28b08-757">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="28b08-757">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="28b08-758">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="28b08-758">Get 500 members maximum.</span></span>
- <span data-ttu-id="28b08-759">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="28b08-759">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="28b08-760">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-760">Type</span></span>

*   <span data-ttu-id="28b08-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="28b08-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-762">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-762">Requirements</span></span>

|<span data-ttu-id="28b08-763">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-763">Requirement</span></span>|<span data-ttu-id="28b08-764">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-764">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-765">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-765">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-766">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-766">1.0</span></span>|
|[<span data-ttu-id="28b08-767">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-767">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-768">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-768">ReadItem</span></span>|
|[<span data-ttu-id="28b08-769">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-769">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-770">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-770">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="28b08-771">Métodos</span><span class="sxs-lookup"><span data-stu-id="28b08-771">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="28b08-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="28b08-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="28b08-773">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="28b08-773">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="28b08-774">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="28b08-774">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="28b08-775">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="28b08-775">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-776">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-776">Parameters</span></span>
|<span data-ttu-id="28b08-777">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-777">Name</span></span>|<span data-ttu-id="28b08-778">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-778">Type</span></span>|<span data-ttu-id="28b08-779">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-779">Attributes</span></span>|<span data-ttu-id="28b08-780">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-780">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="28b08-781">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-781">String</span></span>||<span data-ttu-id="28b08-p144">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="28b08-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="28b08-784">String</span><span class="sxs-lookup"><span data-stu-id="28b08-784">String</span></span>||<span data-ttu-id="28b08-p145">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="28b08-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="28b08-787">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-787">Object</span></span>|<span data-ttu-id="28b08-788">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-788">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-789">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-789">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="28b08-790">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-790">Object</span></span>|<span data-ttu-id="28b08-791">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-791">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-792">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-792">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="28b08-793">Booliano</span><span class="sxs-lookup"><span data-stu-id="28b08-793">Boolean</span></span>|<span data-ttu-id="28b08-794">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-794">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-795">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="28b08-795">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="28b08-796">function</span><span class="sxs-lookup"><span data-stu-id="28b08-796">function</span></span>|<span data-ttu-id="28b08-797">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-797">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-798">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-798">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="28b08-799">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="28b08-799">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="28b08-800">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="28b08-800">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="28b08-801">Erros</span><span class="sxs-lookup"><span data-stu-id="28b08-801">Errors</span></span>

|<span data-ttu-id="28b08-802">Código de erro</span><span class="sxs-lookup"><span data-stu-id="28b08-802">Error code</span></span>|<span data-ttu-id="28b08-803">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-803">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="28b08-804">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="28b08-804">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="28b08-805">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="28b08-805">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="28b08-806">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="28b08-806">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-807">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-807">Requirements</span></span>

|<span data-ttu-id="28b08-808">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-808">Requirement</span></span>|<span data-ttu-id="28b08-809">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-809">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-810">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-810">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-811">1.1</span><span class="sxs-lookup"><span data-stu-id="28b08-811">1.1</span></span>|
|[<span data-ttu-id="28b08-812">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-812">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-813">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="28b08-813">ReadWriteItem</span></span>|
|[<span data-ttu-id="28b08-814">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-814">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-815">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-815">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="28b08-816">Exemplos</span><span class="sxs-lookup"><span data-stu-id="28b08-816">Examples</span></span>

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

<span data-ttu-id="28b08-817">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="28b08-817">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="28b08-818">addFileAttachmentFromBase64Async (base64file, AttachmentName, [Options], [callback])</span><span class="sxs-lookup"><span data-stu-id="28b08-818">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="28b08-819">Adiciona um arquivo da codificação Base64 a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="28b08-819">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="28b08-820">O `addFileAttachmentFromBase64Async` método carrega o arquivo da codificação Base64 e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="28b08-820">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="28b08-821">Esse método retorna o identificador de anexo no objeto AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="28b08-821">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="28b08-822">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="28b08-822">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-823">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-823">Parameters</span></span>

|<span data-ttu-id="28b08-824">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-824">Name</span></span>|<span data-ttu-id="28b08-825">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-825">Type</span></span>|<span data-ttu-id="28b08-826">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-826">Attributes</span></span>|<span data-ttu-id="28b08-827">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-827">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="28b08-828">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-828">String</span></span>||<span data-ttu-id="28b08-829">O conteúdo codificado em Base64 de uma imagem ou arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="28b08-829">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="28b08-830">String</span><span class="sxs-lookup"><span data-stu-id="28b08-830">String</span></span>||<span data-ttu-id="28b08-p147">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="28b08-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="28b08-833">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-833">Object</span></span>|<span data-ttu-id="28b08-834">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-834">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-835">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-835">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="28b08-836">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-836">Object</span></span>|<span data-ttu-id="28b08-837">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-837">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-838">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-838">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="28b08-839">Booliano</span><span class="sxs-lookup"><span data-stu-id="28b08-839">Boolean</span></span>|<span data-ttu-id="28b08-840">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-840">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-841">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="28b08-841">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="28b08-842">function</span><span class="sxs-lookup"><span data-stu-id="28b08-842">function</span></span>|<span data-ttu-id="28b08-843">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-843">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-844">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-844">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="28b08-845">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="28b08-845">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="28b08-846">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="28b08-846">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="28b08-847">Erros</span><span class="sxs-lookup"><span data-stu-id="28b08-847">Errors</span></span>

|<span data-ttu-id="28b08-848">Código de erro</span><span class="sxs-lookup"><span data-stu-id="28b08-848">Error code</span></span>|<span data-ttu-id="28b08-849">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-849">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="28b08-850">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="28b08-850">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="28b08-851">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="28b08-851">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="28b08-852">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="28b08-852">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-853">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-853">Requirements</span></span>

|<span data-ttu-id="28b08-854">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-854">Requirement</span></span>|<span data-ttu-id="28b08-855">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-856">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-857">Visualização</span><span class="sxs-lookup"><span data-stu-id="28b08-857">Preview</span></span>|
|[<span data-ttu-id="28b08-858">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-858">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-859">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="28b08-859">ReadWriteItem</span></span>|
|[<span data-ttu-id="28b08-860">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-860">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-861">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-861">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="28b08-862">Exemplos</span><span class="sxs-lookup"><span data-stu-id="28b08-862">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="28b08-863">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="28b08-863">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="28b08-864">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="28b08-864">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="28b08-865">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="28b08-865">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-866">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-866">Parameters</span></span>

| <span data-ttu-id="28b08-867">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-867">Name</span></span> | <span data-ttu-id="28b08-868">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-868">Type</span></span> | <span data-ttu-id="28b08-869">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-869">Attributes</span></span> | <span data-ttu-id="28b08-870">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-870">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="28b08-871">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="28b08-871">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="28b08-872">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="28b08-872">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="28b08-873">Função</span><span class="sxs-lookup"><span data-stu-id="28b08-873">Function</span></span> || <span data-ttu-id="28b08-p148">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="28b08-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="28b08-877">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-877">Object</span></span> | <span data-ttu-id="28b08-878">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-878">&lt;optional&gt;</span></span> | <span data-ttu-id="28b08-879">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-879">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="28b08-880">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-880">Object</span></span> | <span data-ttu-id="28b08-881">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-881">&lt;optional&gt;</span></span> | <span data-ttu-id="28b08-882">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-882">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="28b08-883">function</span><span class="sxs-lookup"><span data-stu-id="28b08-883">function</span></span>| <span data-ttu-id="28b08-884">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-884">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-885">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-885">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-886">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-886">Requirements</span></span>

|<span data-ttu-id="28b08-887">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-887">Requirement</span></span>| <span data-ttu-id="28b08-888">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-888">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-889">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-889">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="28b08-890">1.7</span><span class="sxs-lookup"><span data-stu-id="28b08-890">1.7</span></span> |
|[<span data-ttu-id="28b08-891">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-891">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="28b08-892">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-892">ReadItem</span></span> |
|[<span data-ttu-id="28b08-893">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-893">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="28b08-894">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-894">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="28b08-895">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-895">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="28b08-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="28b08-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="28b08-897">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-897">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="28b08-p149">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="28b08-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="28b08-901">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="28b08-901">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="28b08-902">Se o Suplemento do Office estiver em execução no Outlook na Web, o método `addItemAttachmentAsync` pode anexar itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="28b08-902">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-903">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-903">Parameters</span></span>

|<span data-ttu-id="28b08-904">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-904">Name</span></span>|<span data-ttu-id="28b08-905">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-905">Type</span></span>|<span data-ttu-id="28b08-906">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-906">Attributes</span></span>|<span data-ttu-id="28b08-907">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-907">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="28b08-908">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-908">String</span></span>||<span data-ttu-id="28b08-p150">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="28b08-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="28b08-911">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-911">String</span></span>||<span data-ttu-id="28b08-912">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="28b08-912">The subject of the item to be attached.</span></span> <span data-ttu-id="28b08-913">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="28b08-913">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="28b08-914">Object</span><span class="sxs-lookup"><span data-stu-id="28b08-914">Object</span></span>|<span data-ttu-id="28b08-915">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-915">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-916">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-916">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="28b08-917">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-917">Object</span></span>|<span data-ttu-id="28b08-918">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-918">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-919">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-919">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="28b08-920">function</span><span class="sxs-lookup"><span data-stu-id="28b08-920">function</span></span>|<span data-ttu-id="28b08-921">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-921">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-922">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-922">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="28b08-923">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="28b08-923">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="28b08-924">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="28b08-924">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="28b08-925">Erros</span><span class="sxs-lookup"><span data-stu-id="28b08-925">Errors</span></span>

|<span data-ttu-id="28b08-926">Código de erro</span><span class="sxs-lookup"><span data-stu-id="28b08-926">Error code</span></span>|<span data-ttu-id="28b08-927">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-927">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="28b08-928">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="28b08-928">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-929">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-929">Requirements</span></span>

|<span data-ttu-id="28b08-930">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-930">Requirement</span></span>|<span data-ttu-id="28b08-931">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-931">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-932">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-932">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-933">1.1</span><span class="sxs-lookup"><span data-stu-id="28b08-933">1.1</span></span>|
|[<span data-ttu-id="28b08-934">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-934">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-935">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="28b08-935">ReadWriteItem</span></span>|
|[<span data-ttu-id="28b08-936">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-936">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-937">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-937">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-938">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-938">Example</span></span>

<span data-ttu-id="28b08-939">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="28b08-939">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="28b08-940">close()</span><span class="sxs-lookup"><span data-stu-id="28b08-940">close()</span></span>

<span data-ttu-id="28b08-941">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="28b08-941">Closes the current item that is being composed.</span></span>

<span data-ttu-id="28b08-p152">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="28b08-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-944">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="28b08-944">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="28b08-945">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="28b08-945">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-946">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-946">Requirements</span></span>

|<span data-ttu-id="28b08-947">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-947">Requirement</span></span>|<span data-ttu-id="28b08-948">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-948">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-949">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-949">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-950">1.3</span><span class="sxs-lookup"><span data-stu-id="28b08-950">1.3</span></span>|
|[<span data-ttu-id="28b08-951">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-951">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-952">Restrito</span><span class="sxs-lookup"><span data-stu-id="28b08-952">Restricted</span></span>|
|[<span data-ttu-id="28b08-953">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-953">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-954">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-954">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="28b08-955">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="28b08-955">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="28b08-956">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="28b08-956">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-957">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="28b08-957">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="28b08-958">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="28b08-958">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="28b08-959">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="28b08-959">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="28b08-p153">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="28b08-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-963">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-963">Parameters</span></span>

|<span data-ttu-id="28b08-964">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-964">Name</span></span>|<span data-ttu-id="28b08-965">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-965">Type</span></span>|<span data-ttu-id="28b08-966">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-966">Attributes</span></span>|<span data-ttu-id="28b08-967">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-967">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="28b08-968">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="28b08-968">String &#124; Object</span></span>||<span data-ttu-id="28b08-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="28b08-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="28b08-971">**OU**</span><span class="sxs-lookup"><span data-stu-id="28b08-971">**OR**</span></span><br/><span data-ttu-id="28b08-p155">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="28b08-974">Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-974">String</span></span>|<span data-ttu-id="28b08-975">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-975">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-p156">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="28b08-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="28b08-978">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-978">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="28b08-979">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-979">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-980">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="28b08-980">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="28b08-981">Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-981">String</span></span>||<span data-ttu-id="28b08-p157">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="28b08-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="28b08-984">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-984">String</span></span>||<span data-ttu-id="28b08-985">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="28b08-985">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="28b08-986">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-986">String</span></span>||<span data-ttu-id="28b08-p158">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="28b08-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="28b08-989">Booliano</span><span class="sxs-lookup"><span data-stu-id="28b08-989">Boolean</span></span>||<span data-ttu-id="28b08-p159">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="28b08-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="28b08-992">String</span><span class="sxs-lookup"><span data-stu-id="28b08-992">String</span></span>||<span data-ttu-id="28b08-p160">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="28b08-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="28b08-996">function</span><span class="sxs-lookup"><span data-stu-id="28b08-996">function</span></span>|<span data-ttu-id="28b08-997">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-997">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-998">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-998">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-999">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-999">Requirements</span></span>

|<span data-ttu-id="28b08-1000">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1000">Requirement</span></span>|<span data-ttu-id="28b08-1001">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1002">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-1003">1.0</span></span>|
|[<span data-ttu-id="28b08-1004">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1004">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1005">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1006">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-1006">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1007">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-1007">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="28b08-1008">Exemplos</span><span class="sxs-lookup"><span data-stu-id="28b08-1008">Examples</span></span>

<span data-ttu-id="28b08-1009">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="28b08-1009">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="28b08-1010">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="28b08-1010">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="28b08-1011">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="28b08-1011">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="28b08-1012">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="28b08-1012">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="28b08-1013">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="28b08-1013">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="28b08-1014">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1014">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="28b08-1015">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="28b08-1015">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="28b08-1016">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="28b08-1016">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-1017">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="28b08-1017">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="28b08-1018">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="28b08-1018">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="28b08-1019">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="28b08-1019">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="28b08-p161">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="28b08-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1023">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1023">Parameters</span></span>

|<span data-ttu-id="28b08-1024">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1024">Name</span></span>|<span data-ttu-id="28b08-1025">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1025">Type</span></span>|<span data-ttu-id="28b08-1026">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-1026">Attributes</span></span>|<span data-ttu-id="28b08-1027">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1027">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="28b08-1028">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="28b08-1028">String &#124; Object</span></span>||<span data-ttu-id="28b08-p162">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="28b08-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="28b08-1031">**OU**</span><span class="sxs-lookup"><span data-stu-id="28b08-1031">**OR**</span></span><br/><span data-ttu-id="28b08-p163">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="28b08-1034">Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-1034">String</span></span>|<span data-ttu-id="28b08-1035">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1035">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-p164">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="28b08-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="28b08-1038">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1038">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="28b08-1039">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1040">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="28b08-1040">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="28b08-1041">Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-1041">String</span></span>||<span data-ttu-id="28b08-p165">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="28b08-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="28b08-1044">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-1044">String</span></span>||<span data-ttu-id="28b08-1045">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="28b08-1045">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="28b08-1046">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-1046">String</span></span>||<span data-ttu-id="28b08-p166">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="28b08-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="28b08-1049">Booliano</span><span class="sxs-lookup"><span data-stu-id="28b08-1049">Boolean</span></span>||<span data-ttu-id="28b08-p167">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="28b08-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="28b08-1052">String</span><span class="sxs-lookup"><span data-stu-id="28b08-1052">String</span></span>||<span data-ttu-id="28b08-p168">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="28b08-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="28b08-1056">function</span><span class="sxs-lookup"><span data-stu-id="28b08-1056">function</span></span>|<span data-ttu-id="28b08-1057">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1058">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1059">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1059">Requirements</span></span>

|<span data-ttu-id="28b08-1060">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1060">Requirement</span></span>|<span data-ttu-id="28b08-1061">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1061">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1062">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1062">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1063">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-1063">1.0</span></span>|
|[<span data-ttu-id="28b08-1064">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1064">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1065">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1065">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1066">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-1066">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1067">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-1067">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="28b08-1068">Exemplos</span><span class="sxs-lookup"><span data-stu-id="28b08-1068">Examples</span></span>

<span data-ttu-id="28b08-1069">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="28b08-1069">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="28b08-1070">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="28b08-1070">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="28b08-1071">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="28b08-1071">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="28b08-1072">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="28b08-1072">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="28b08-1073">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="28b08-1073">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="28b08-1074">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1074">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="28b08-1075">getAllInternetHeadersAsync ([opções], [callback])</span><span class="sxs-lookup"><span data-stu-id="28b08-1075">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="28b08-1076">Obtém todos os cabeçalhos de Internet da mensagem como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="28b08-1076">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="28b08-1077">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="28b08-1077">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1078">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1078">Parameters</span></span>

|<span data-ttu-id="28b08-1079">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1079">Name</span></span>|<span data-ttu-id="28b08-1080">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1080">Type</span></span>|<span data-ttu-id="28b08-1081">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-1081">Attributes</span></span>|<span data-ttu-id="28b08-1082">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1082">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="28b08-1083">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1083">Object</span></span>|<span data-ttu-id="28b08-1084">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1084">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1085">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-1085">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="28b08-1086">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1086">Object</span></span>|<span data-ttu-id="28b08-1087">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1087">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1088">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1088">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="28b08-1089">function</span><span class="sxs-lookup"><span data-stu-id="28b08-1089">function</span></span>|<span data-ttu-id="28b08-1090">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1091">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-1091">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="28b08-1092">Com êxito, os dados de cabeçalhos de Internet são fornecidos na propriedade asyncResult. Value como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="28b08-1092">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="28b08-1093">Consulte [RFC 2183](https://tools.ietf.org/html/rfc2183) para obter as informações de formatação do valor de cadeia de caracteres retornado.</span><span class="sxs-lookup"><span data-stu-id="28b08-1093">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="28b08-1094">Se a chamada falhar, a propriedade asyncResult. Error conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="28b08-1094">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1095">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1095">Requirements</span></span>

|<span data-ttu-id="28b08-1096">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1096">Requirement</span></span>|<span data-ttu-id="28b08-1097">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1098">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1099">Visualização</span><span class="sxs-lookup"><span data-stu-id="28b08-1099">Preview</span></span>|
|[<span data-ttu-id="28b08-1100">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1101">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1102">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1103">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-1103">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="28b08-1104">Retorna:</span><span class="sxs-lookup"><span data-stu-id="28b08-1104">Returns:</span></span>

<span data-ttu-id="28b08-1105">A Internet cabeçalhos dados como uma cadeia de caracteres formatada de acordo com a [RFC 2183](https://tools.ietf.org/html/rfc2183).</span><span class="sxs-lookup"><span data-stu-id="28b08-1105">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="28b08-1106">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="28b08-1106">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="28b08-1107">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1107">Example</span></span>

```js
// Get the internet headers related to the mail.
Office.context.mailbox.item.getAllInternetHeadersAsync(
  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(asyncResult.value);
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="28b08-1108">getAttachmentContentAsync (attachmentid, [opções], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="28b08-1108">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="28b08-1109">Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um `AttachmentContent` objeto.</span><span class="sxs-lookup"><span data-stu-id="28b08-1109">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="28b08-1110">O `getAttachmentContentAsync` método obtém o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="28b08-1110">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="28b08-1111">Como prática recomendada, você deve usar o identificador para recuperar um anexo na mesma sessão em que o attachmentIds foi recuperado com a `getAttachmentsAsync` chamada ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="28b08-1111">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="28b08-1112">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="28b08-1112">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="28b08-1113">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1113">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1114">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1114">Parameters</span></span>

|<span data-ttu-id="28b08-1115">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1115">Name</span></span>|<span data-ttu-id="28b08-1116">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1116">Type</span></span>|<span data-ttu-id="28b08-1117">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-1117">Attributes</span></span>|<span data-ttu-id="28b08-1118">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1118">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="28b08-1119">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-1119">String</span></span>||<span data-ttu-id="28b08-1120">O identificador do anexo que você deseja obter.</span><span class="sxs-lookup"><span data-stu-id="28b08-1120">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="28b08-1121">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1121">Object</span></span>|<span data-ttu-id="28b08-1122">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1122">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1123">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-1123">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="28b08-1124">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1124">Object</span></span>|<span data-ttu-id="28b08-1125">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1125">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1126">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1126">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="28b08-1127">function</span><span class="sxs-lookup"><span data-stu-id="28b08-1127">function</span></span>|<span data-ttu-id="28b08-1128">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1129">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-1129">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1130">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1130">Requirements</span></span>

|<span data-ttu-id="28b08-1131">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1131">Requirement</span></span>|<span data-ttu-id="28b08-1132">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1133">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1134">Visualização</span><span class="sxs-lookup"><span data-stu-id="28b08-1134">Preview</span></span>|
|[<span data-ttu-id="28b08-1135">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1136">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1137">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-1137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1138">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-1138">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="28b08-1139">Retorna:</span><span class="sxs-lookup"><span data-stu-id="28b08-1139">Returns:</span></span>

<span data-ttu-id="28b08-1140">Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="28b08-1140">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="28b08-1141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1141">Example</span></span>

```js
var item = Office.context.mailbox.item;
var listOfAttachments = [];
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

<br>

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="28b08-1142">getAttachmentsAsync ([Options], [callback]) → array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="28b08-1142">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="28b08-1143">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="28b08-1143">Gets the item's attachments as an array.</span></span> <span data-ttu-id="28b08-1144">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="28b08-1144">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1145">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1145">Parameters</span></span>

|<span data-ttu-id="28b08-1146">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1146">Name</span></span>|<span data-ttu-id="28b08-1147">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1147">Type</span></span>|<span data-ttu-id="28b08-1148">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-1148">Attributes</span></span>|<span data-ttu-id="28b08-1149">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1149">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="28b08-1150">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1150">Object</span></span>|<span data-ttu-id="28b08-1151">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1151">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1152">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-1152">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="28b08-1153">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1153">Object</span></span>|<span data-ttu-id="28b08-1154">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1155">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1155">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="28b08-1156">function</span><span class="sxs-lookup"><span data-stu-id="28b08-1156">function</span></span>|<span data-ttu-id="28b08-1157">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1158">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-1158">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1159">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1159">Requirements</span></span>

|<span data-ttu-id="28b08-1160">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1160">Requirement</span></span>|<span data-ttu-id="28b08-1161">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1162">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1163">Visualização</span><span class="sxs-lookup"><span data-stu-id="28b08-1163">Preview</span></span>|
|[<span data-ttu-id="28b08-1164">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1165">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1165">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1166">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1167">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-1167">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="28b08-1168">Retorna:</span><span class="sxs-lookup"><span data-stu-id="28b08-1168">Returns:</span></span>

<span data-ttu-id="28b08-1169">Tipo: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="28b08-1169">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="28b08-1170">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1170">Example</span></span>

<span data-ttu-id="28b08-1171">O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="28b08-1171">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="28b08-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="28b08-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="28b08-1173">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="28b08-1173">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-1174">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="28b08-1174">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-1175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1175">Requirements</span></span>

|<span data-ttu-id="28b08-1176">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1176">Requirement</span></span>|<span data-ttu-id="28b08-1177">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1177">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1179">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-1179">1.0</span></span>|
|[<span data-ttu-id="28b08-1180">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1180">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1181">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1182">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-1182">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1183">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-1183">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="28b08-1184">Retorna:</span><span class="sxs-lookup"><span data-stu-id="28b08-1184">Returns:</span></span>

<span data-ttu-id="28b08-1185">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="28b08-1185">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="28b08-1186">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1186">Example</span></span>

<span data-ttu-id="28b08-1187">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="28b08-1187">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="28b08-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="28b08-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="28b08-1189">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="28b08-1189">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-1190">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="28b08-1190">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1191">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1191">Parameters</span></span>

|<span data-ttu-id="28b08-1192">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1192">Name</span></span>|<span data-ttu-id="28b08-1193">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1193">Type</span></span>|<span data-ttu-id="28b08-1194">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1194">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="28b08-1195">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="28b08-1195">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="28b08-1196">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="28b08-1196">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1197">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1197">Requirements</span></span>

|<span data-ttu-id="28b08-1198">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1198">Requirement</span></span>|<span data-ttu-id="28b08-1199">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1199">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1200">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1201">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-1201">1.0</span></span>|
|[<span data-ttu-id="28b08-1202">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1203">Restrito</span><span class="sxs-lookup"><span data-stu-id="28b08-1203">Restricted</span></span>|
|[<span data-ttu-id="28b08-1204">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-1204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1205">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-1205">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="28b08-1206">Retorna:</span><span class="sxs-lookup"><span data-stu-id="28b08-1206">Returns:</span></span>

<span data-ttu-id="28b08-1207">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="28b08-1207">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="28b08-1208">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="28b08-1208">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="28b08-1209">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="28b08-1209">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="28b08-1210">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-1210">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="28b08-1211">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="28b08-1211">Value of `entityType`</span></span>|<span data-ttu-id="28b08-1212">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="28b08-1212">Type of objects in returned array</span></span>|<span data-ttu-id="28b08-1213">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="28b08-1213">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="28b08-1214">String</span><span class="sxs-lookup"><span data-stu-id="28b08-1214">String</span></span>|<span data-ttu-id="28b08-1215">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="28b08-1215">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="28b08-1216">Contato</span><span class="sxs-lookup"><span data-stu-id="28b08-1216">Contact</span></span>|<span data-ttu-id="28b08-1217">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="28b08-1217">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="28b08-1218">String</span><span class="sxs-lookup"><span data-stu-id="28b08-1218">String</span></span>|<span data-ttu-id="28b08-1219">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="28b08-1219">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="28b08-1220">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="28b08-1220">MeetingSuggestion</span></span>|<span data-ttu-id="28b08-1221">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="28b08-1221">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="28b08-1222">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="28b08-1222">PhoneNumber</span></span>|<span data-ttu-id="28b08-1223">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="28b08-1223">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="28b08-1224">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="28b08-1224">TaskSuggestion</span></span>|<span data-ttu-id="28b08-1225">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="28b08-1225">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="28b08-1226">String</span><span class="sxs-lookup"><span data-stu-id="28b08-1226">String</span></span>|<span data-ttu-id="28b08-1227">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="28b08-1227">**Restricted**</span></span>|

<span data-ttu-id="28b08-1228">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="28b08-1228">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="28b08-1229">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1229">Example</span></span>

<span data-ttu-id="28b08-1230">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="28b08-1230">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="28b08-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="28b08-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="28b08-1232">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="28b08-1232">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-1233">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="28b08-1233">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="28b08-1234">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="28b08-1234">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1235">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1235">Parameters</span></span>

|<span data-ttu-id="28b08-1236">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1236">Name</span></span>|<span data-ttu-id="28b08-1237">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1237">Type</span></span>|<span data-ttu-id="28b08-1238">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1238">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="28b08-1239">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28b08-1239">String</span></span>|<span data-ttu-id="28b08-1240">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="28b08-1240">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1241">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1241">Requirements</span></span>

|<span data-ttu-id="28b08-1242">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1242">Requirement</span></span>|<span data-ttu-id="28b08-1243">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1243">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1244">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1244">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1245">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-1245">1.0</span></span>|
|[<span data-ttu-id="28b08-1246">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1246">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1247">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1247">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1248">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-1248">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1249">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-1249">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="28b08-1250">Retorna:</span><span class="sxs-lookup"><span data-stu-id="28b08-1250">Returns:</span></span>

<span data-ttu-id="28b08-p174">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="28b08-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="28b08-1253">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="28b08-1253">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="28b08-1254">getInitializationContextAsync ([opções], [callback])</span><span class="sxs-lookup"><span data-stu-id="28b08-1254">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="28b08-1255">Obtém dados de inicialização passados quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="28b08-1255">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-1256">Este método só é compatível com o Outlook 2016 ou posterior no Windows (clique para executar versões posteriores a 16.0.8413.1000) e Outlook na Web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="28b08-1256">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1257">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1257">Parameters</span></span>

|<span data-ttu-id="28b08-1258">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1258">Name</span></span>|<span data-ttu-id="28b08-1259">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1259">Type</span></span>|<span data-ttu-id="28b08-1260">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-1260">Attributes</span></span>|<span data-ttu-id="28b08-1261">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1261">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="28b08-1262">Object</span><span class="sxs-lookup"><span data-stu-id="28b08-1262">Object</span></span>|<span data-ttu-id="28b08-1263">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1263">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1264">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-1264">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="28b08-1265">Object</span><span class="sxs-lookup"><span data-stu-id="28b08-1265">Object</span></span>|<span data-ttu-id="28b08-1266">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1266">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1267">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1267">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="28b08-1268">function</span><span class="sxs-lookup"><span data-stu-id="28b08-1268">function</span></span>|<span data-ttu-id="28b08-1269">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1269">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1270">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-1270">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="28b08-1271">Com êxito, os dados de inicialização são fornecidos na `asyncResult.value` Propriedade como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="28b08-1271">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="28b08-1272">Se não houver nenhum contexto de inicialização, `asyncResult` o objeto conterá `Error` um objeto com `code` sua propriedade definida `9020` como e `name` sua propriedade definida `GenericResponseError`como.</span><span class="sxs-lookup"><span data-stu-id="28b08-1272">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1273">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1273">Requirements</span></span>

|<span data-ttu-id="28b08-1274">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1274">Requirement</span></span>|<span data-ttu-id="28b08-1275">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1275">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1276">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1277">Visualização</span><span class="sxs-lookup"><span data-stu-id="28b08-1277">Preview</span></span>|
|[<span data-ttu-id="28b08-1278">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1279">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1280">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-1280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1281">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-1281">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-1282">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1282">Example</span></span>

```js
// Get the initialization context (if present).
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object.
        var context = JSON.parse(asyncResult.value);
        // Do something with context.
      } else {
        // Empty context, treat as no context.
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="28b08-1283">getItemIdAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="28b08-1283">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="28b08-1284">Obtém de forma assíncrona a ID de um item salvo.</span><span class="sxs-lookup"><span data-stu-id="28b08-1284">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="28b08-1285">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="28b08-1285">Compose mode only.</span></span>

<span data-ttu-id="28b08-1286">Quando invocado, este método retorna a ID do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1286">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-1287">Se seu suplemento chamar `getItemIdAsync` um item no modo de redação (por exemplo, para `itemId` usar com o EWS ou a API REST), lembre-se de que, quando o Outlook estiver no modo cache, pode levar algum tempo para que o item seja sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="28b08-1287">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="28b08-1288">Até que o item seja sincronizado, `itemId` o não é reconhecido e usado retorna um erro.</span><span class="sxs-lookup"><span data-stu-id="28b08-1288">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1289">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1289">Parameters</span></span>

|<span data-ttu-id="28b08-1290">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1290">Name</span></span>|<span data-ttu-id="28b08-1291">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1291">Type</span></span>|<span data-ttu-id="28b08-1292">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-1292">Attributes</span></span>|<span data-ttu-id="28b08-1293">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1293">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="28b08-1294">Object</span><span class="sxs-lookup"><span data-stu-id="28b08-1294">Object</span></span>|<span data-ttu-id="28b08-1295">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1295">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1296">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-1296">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="28b08-1297">Object</span><span class="sxs-lookup"><span data-stu-id="28b08-1297">Object</span></span>|<span data-ttu-id="28b08-1298">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1298">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1299">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1299">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="28b08-1300">function</span><span class="sxs-lookup"><span data-stu-id="28b08-1300">function</span></span>||<span data-ttu-id="28b08-1301">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-1301">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="28b08-1302">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="28b08-1302">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="28b08-1303">Erros</span><span class="sxs-lookup"><span data-stu-id="28b08-1303">Errors</span></span>

|<span data-ttu-id="28b08-1304">Código de erro</span><span class="sxs-lookup"><span data-stu-id="28b08-1304">Error code</span></span>|<span data-ttu-id="28b08-1305">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1305">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="28b08-1306">A ID não pode ser recuperada até que o item seja salvo.</span><span class="sxs-lookup"><span data-stu-id="28b08-1306">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1307">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1307">Requirements</span></span>

|<span data-ttu-id="28b08-1308">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1308">Requirement</span></span>|<span data-ttu-id="28b08-1309">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1309">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1310">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1310">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1311">Visualização</span><span class="sxs-lookup"><span data-stu-id="28b08-1311">Preview</span></span>|
|[<span data-ttu-id="28b08-1312">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1312">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1313">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1314">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-1314">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1315">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-1315">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="28b08-1316">Exemplos</span><span class="sxs-lookup"><span data-stu-id="28b08-1316">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="28b08-1317">O exemplo a seguir mostra a estrutura do `result` parâmetro que é passado para a função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1317">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="28b08-1318">A `value` propriedade contém a ID do item.</span><span class="sxs-lookup"><span data-stu-id="28b08-1318">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="28b08-1319">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="28b08-1319">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="28b08-1320">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="28b08-1320">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-1321">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="28b08-1321">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="28b08-p178">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="28b08-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="28b08-1325">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="28b08-1325">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="28b08-1326">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="28b08-1326">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="28b08-p179">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="28b08-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-1330">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1330">Requirements</span></span>

|<span data-ttu-id="28b08-1331">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1331">Requirement</span></span>|<span data-ttu-id="28b08-1332">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1332">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1333">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1334">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-1334">1.0</span></span>|
|[<span data-ttu-id="28b08-1335">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1336">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1337">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-1337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1338">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-1338">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="28b08-1339">Retorna:</span><span class="sxs-lookup"><span data-stu-id="28b08-1339">Returns:</span></span>

<span data-ttu-id="28b08-p180">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="28b08-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="28b08-1342">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="28b08-1342">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="28b08-1343">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1343">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="28b08-1344">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1344">Example</span></span>

<span data-ttu-id="28b08-1345">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="28b08-1345">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="28b08-1346">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="28b08-1346">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="28b08-1347">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="28b08-1347">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-1348">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="28b08-1348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="28b08-1349">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="28b08-1349">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="28b08-p181">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="28b08-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1352">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1352">Parameters</span></span>

|<span data-ttu-id="28b08-1353">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1353">Name</span></span>|<span data-ttu-id="28b08-1354">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1354">Type</span></span>|<span data-ttu-id="28b08-1355">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1355">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="28b08-1356">String</span><span class="sxs-lookup"><span data-stu-id="28b08-1356">String</span></span>|<span data-ttu-id="28b08-1357">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="28b08-1357">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1358">Requirements</span></span>

|<span data-ttu-id="28b08-1359">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1359">Requirement</span></span>|<span data-ttu-id="28b08-1360">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1360">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1361">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1362">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-1362">1.0</span></span>|
|[<span data-ttu-id="28b08-1363">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1364">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1365">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-1365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1366">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-1366">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="28b08-1367">Retorna:</span><span class="sxs-lookup"><span data-stu-id="28b08-1367">Returns:</span></span>

<span data-ttu-id="28b08-1368">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="28b08-1368">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="28b08-1369">Tipo: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="28b08-1369">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="28b08-1370">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1370">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="28b08-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="28b08-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="28b08-1372">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="28b08-1372">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="28b08-p182">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="28b08-p182">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1375">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1375">Parameters</span></span>

|<span data-ttu-id="28b08-1376">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1376">Name</span></span>|<span data-ttu-id="28b08-1377">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1377">Type</span></span>|<span data-ttu-id="28b08-1378">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-1378">Attributes</span></span>|<span data-ttu-id="28b08-1379">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1379">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="28b08-1380">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="28b08-1380">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="28b08-p183">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="28b08-p183">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="28b08-1384">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1384">Object</span></span>|<span data-ttu-id="28b08-1385">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1385">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1386">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-1386">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="28b08-1387">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1387">Object</span></span>|<span data-ttu-id="28b08-1388">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1388">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1389">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1389">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="28b08-1390">function</span><span class="sxs-lookup"><span data-stu-id="28b08-1390">function</span></span>||<span data-ttu-id="28b08-1391">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-1391">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="28b08-1392">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="28b08-1392">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="28b08-1393">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="28b08-1393">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1394">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1394">Requirements</span></span>

|<span data-ttu-id="28b08-1395">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1395">Requirement</span></span>|<span data-ttu-id="28b08-1396">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1396">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1397">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1398">1.2</span><span class="sxs-lookup"><span data-stu-id="28b08-1398">1.2</span></span>|
|[<span data-ttu-id="28b08-1399">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1400">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1401">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-1401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1402">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-1402">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="28b08-1403">Retorna:</span><span class="sxs-lookup"><span data-stu-id="28b08-1403">Returns:</span></span>

<span data-ttu-id="28b08-1404">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="28b08-1404">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="28b08-1405">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="28b08-1405">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="28b08-1406">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1406">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="28b08-1407">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="28b08-1407">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="28b08-1408">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="28b08-1408">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="28b08-1409">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="28b08-1409">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-1410">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="28b08-1410">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-1411">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1411">Requirements</span></span>

|<span data-ttu-id="28b08-1412">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1412">Requirement</span></span>|<span data-ttu-id="28b08-1413">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1413">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1414">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1414">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1415">1.6</span><span class="sxs-lookup"><span data-stu-id="28b08-1415">1.6</span></span>|
|[<span data-ttu-id="28b08-1416">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1417">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1418">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-1418">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1419">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-1419">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="28b08-1420">Retorna:</span><span class="sxs-lookup"><span data-stu-id="28b08-1420">Returns:</span></span>

<span data-ttu-id="28b08-1421">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="28b08-1421">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="28b08-1422">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1422">Example</span></span>

<span data-ttu-id="28b08-1423">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="28b08-1423">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="28b08-1424">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="28b08-1424">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="28b08-p186">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="28b08-p186">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-1427">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="28b08-1427">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="28b08-p187">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="28b08-p187">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="28b08-1431">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="28b08-1431">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="28b08-1432">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="28b08-1432">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="28b08-p188">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="28b08-p188">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="28b08-1436">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1436">Requirements</span></span>

|<span data-ttu-id="28b08-1437">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1437">Requirement</span></span>|<span data-ttu-id="28b08-1438">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1438">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1439">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1440">1.6</span><span class="sxs-lookup"><span data-stu-id="28b08-1440">1.6</span></span>|
|[<span data-ttu-id="28b08-1441">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1441">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1442">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1443">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-1443">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1444">Read</span><span class="sxs-lookup"><span data-stu-id="28b08-1444">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="28b08-1445">Retorna:</span><span class="sxs-lookup"><span data-stu-id="28b08-1445">Returns:</span></span>

<span data-ttu-id="28b08-p189">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="28b08-p189">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="28b08-1448">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1448">Example</span></span>

<span data-ttu-id="28b08-1449">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="28b08-1449">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="28b08-1450">getSharedPropertiesAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="28b08-1450">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="28b08-1451">Obtém as propriedades do compromisso ou da mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="28b08-1451">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1452">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1452">Parameters</span></span>

|<span data-ttu-id="28b08-1453">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1453">Name</span></span>|<span data-ttu-id="28b08-1454">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1454">Type</span></span>|<span data-ttu-id="28b08-1455">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-1455">Attributes</span></span>|<span data-ttu-id="28b08-1456">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1456">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="28b08-1457">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1457">Object</span></span>|<span data-ttu-id="28b08-1458">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1458">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1459">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-1459">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="28b08-1460">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1460">Object</span></span>|<span data-ttu-id="28b08-1461">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1461">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1462">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1462">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="28b08-1463">function</span><span class="sxs-lookup"><span data-stu-id="28b08-1463">function</span></span>||<span data-ttu-id="28b08-1464">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-1464">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="28b08-1465">As propriedades compartilhadas são fornecidas [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) como um objeto `asyncResult.value` na propriedade.</span><span class="sxs-lookup"><span data-stu-id="28b08-1465">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="28b08-1466">Este objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="28b08-1466">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1467">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1467">Requirements</span></span>

|<span data-ttu-id="28b08-1468">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1468">Requirement</span></span>|<span data-ttu-id="28b08-1469">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1469">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1470">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1471">Visualização</span><span class="sxs-lookup"><span data-stu-id="28b08-1471">Preview</span></span>|
|[<span data-ttu-id="28b08-1472">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1473">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1474">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-1474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1475">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-1475">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-1476">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1476">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="28b08-1477">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="28b08-1477">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="28b08-1478">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="28b08-1478">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="28b08-p191">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="28b08-p191">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1482">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1482">Parameters</span></span>

|<span data-ttu-id="28b08-1483">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1483">Name</span></span>|<span data-ttu-id="28b08-1484">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1484">Type</span></span>|<span data-ttu-id="28b08-1485">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-1485">Attributes</span></span>|<span data-ttu-id="28b08-1486">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1486">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="28b08-1487">function</span><span class="sxs-lookup"><span data-stu-id="28b08-1487">function</span></span>||<span data-ttu-id="28b08-1488">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-1488">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="28b08-1489">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="28b08-1489">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="28b08-1490">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="28b08-1490">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="28b08-1491">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1491">Object</span></span>|<span data-ttu-id="28b08-1492">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1492">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1493">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1493">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="28b08-1494">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1494">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1495">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1495">Requirements</span></span>

|<span data-ttu-id="28b08-1496">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1496">Requirement</span></span>|<span data-ttu-id="28b08-1497">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1497">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1498">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1498">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1499">1.0</span><span class="sxs-lookup"><span data-stu-id="28b08-1499">1.0</span></span>|
|[<span data-ttu-id="28b08-1500">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1500">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1501">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1501">ReadItem</span></span>|
|[<span data-ttu-id="28b08-1502">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-1502">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1503">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-1503">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-1504">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1504">Example</span></span>

<span data-ttu-id="28b08-p194">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="28b08-p194">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="28b08-1508">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="28b08-1508">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="28b08-1509">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="28b08-1509">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="28b08-1510">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="28b08-1510">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="28b08-1511">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="28b08-1511">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="28b08-1512">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="28b08-1512">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="28b08-1513">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1513">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1514">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1514">Parameters</span></span>

|<span data-ttu-id="28b08-1515">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1515">Name</span></span>|<span data-ttu-id="28b08-1516">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1516">Type</span></span>|<span data-ttu-id="28b08-1517">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-1517">Attributes</span></span>|<span data-ttu-id="28b08-1518">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1518">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="28b08-1519">String</span><span class="sxs-lookup"><span data-stu-id="28b08-1519">String</span></span>||<span data-ttu-id="28b08-1520">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="28b08-1520">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="28b08-1521">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1521">Object</span></span>|<span data-ttu-id="28b08-1522">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1522">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1523">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-1523">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="28b08-1524">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1524">Object</span></span>|<span data-ttu-id="28b08-1525">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1525">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1526">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1526">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="28b08-1527">function</span><span class="sxs-lookup"><span data-stu-id="28b08-1527">function</span></span>|<span data-ttu-id="28b08-1528">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1528">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1529">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-1529">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="28b08-1530">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="28b08-1530">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="28b08-1531">Erros</span><span class="sxs-lookup"><span data-stu-id="28b08-1531">Errors</span></span>

|<span data-ttu-id="28b08-1532">Código de erro</span><span class="sxs-lookup"><span data-stu-id="28b08-1532">Error code</span></span>|<span data-ttu-id="28b08-1533">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1533">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="28b08-1534">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="28b08-1534">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1535">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1535">Requirements</span></span>

|<span data-ttu-id="28b08-1536">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1536">Requirement</span></span>|<span data-ttu-id="28b08-1537">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1537">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1538">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1539">1.1</span><span class="sxs-lookup"><span data-stu-id="28b08-1539">1.1</span></span>|
|[<span data-ttu-id="28b08-1540">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1541">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1541">ReadWriteItem</span></span>|
|[<span data-ttu-id="28b08-1542">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-1542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1543">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-1543">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-1544">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1544">Example</span></span>

<span data-ttu-id="28b08-1545">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="28b08-1545">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="28b08-1546">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="28b08-1546">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="28b08-1547">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="28b08-1547">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="28b08-1548">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="28b08-1548">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1549">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1549">Parameters</span></span>

| <span data-ttu-id="28b08-1550">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1550">Name</span></span> | <span data-ttu-id="28b08-1551">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1551">Type</span></span> | <span data-ttu-id="28b08-1552">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-1552">Attributes</span></span> | <span data-ttu-id="28b08-1553">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1553">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="28b08-1554">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="28b08-1554">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="28b08-1555">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="28b08-1555">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="28b08-1556">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1556">Object</span></span> | <span data-ttu-id="28b08-1557">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1557">&lt;optional&gt;</span></span> | <span data-ttu-id="28b08-1558">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-1558">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="28b08-1559">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1559">Object</span></span> | <span data-ttu-id="28b08-1560">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1560">&lt;optional&gt;</span></span> | <span data-ttu-id="28b08-1561">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1561">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="28b08-1562">function</span><span class="sxs-lookup"><span data-stu-id="28b08-1562">function</span></span>| <span data-ttu-id="28b08-1563">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1563">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1564">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-1564">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1565">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1565">Requirements</span></span>

|<span data-ttu-id="28b08-1566">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1566">Requirement</span></span>| <span data-ttu-id="28b08-1567">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1567">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1568">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1568">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="28b08-1569">1.7</span><span class="sxs-lookup"><span data-stu-id="28b08-1569">1.7</span></span> |
|[<span data-ttu-id="28b08-1570">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1570">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="28b08-1571">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1571">ReadItem</span></span> |
|[<span data-ttu-id="28b08-1572">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="28b08-1572">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="28b08-1573">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28b08-1573">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="28b08-1574">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="28b08-1574">saveAsync([options], callback)</span></span>

<span data-ttu-id="28b08-1575">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="28b08-1575">Asynchronously saves an item.</span></span>

<span data-ttu-id="28b08-1576">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1576">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="28b08-1577">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="28b08-1577">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="28b08-1578">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="28b08-1578">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-1579">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="28b08-1579">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="28b08-1580">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="28b08-1580">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="28b08-p198">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="28b08-p198">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="28b08-1584">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="28b08-1584">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="28b08-1585">O Outlook no Mac não dá suporte ao salvamento de reuniões.</span><span class="sxs-lookup"><span data-stu-id="28b08-1585">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="28b08-1586">O método `saveAsync` falha quando chamado a partir de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="28b08-1586">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="28b08-1587">Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="28b08-1587">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="28b08-1588">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="28b08-1588">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1589">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1589">Parameters</span></span>

|<span data-ttu-id="28b08-1590">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1590">Name</span></span>|<span data-ttu-id="28b08-1591">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1591">Type</span></span>|<span data-ttu-id="28b08-1592">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-1592">Attributes</span></span>|<span data-ttu-id="28b08-1593">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1593">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="28b08-1594">Object</span><span class="sxs-lookup"><span data-stu-id="28b08-1594">Object</span></span>|<span data-ttu-id="28b08-1595">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1595">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1596">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-1596">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="28b08-1597">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1597">Object</span></span>|<span data-ttu-id="28b08-1598">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1598">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1599">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1599">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="28b08-1600">function</span><span class="sxs-lookup"><span data-stu-id="28b08-1600">function</span></span>||<span data-ttu-id="28b08-1601">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-1601">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="28b08-1602">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="28b08-1602">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1603">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1603">Requirements</span></span>

|<span data-ttu-id="28b08-1604">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1604">Requirement</span></span>|<span data-ttu-id="28b08-1605">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1605">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1606">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1607">1.3</span><span class="sxs-lookup"><span data-stu-id="28b08-1607">1.3</span></span>|
|[<span data-ttu-id="28b08-1608">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1609">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1609">ReadWriteItem</span></span>|
|[<span data-ttu-id="28b08-1610">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-1610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1611">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-1611">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="28b08-1612">Exemplos</span><span class="sxs-lookup"><span data-stu-id="28b08-1612">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="28b08-p200">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="28b08-p200">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="28b08-1615">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="28b08-1615">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="28b08-1616">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="28b08-1616">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="28b08-p201">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="28b08-p201">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="28b08-1620">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="28b08-1620">Parameters</span></span>

|<span data-ttu-id="28b08-1621">Nome</span><span class="sxs-lookup"><span data-stu-id="28b08-1621">Name</span></span>|<span data-ttu-id="28b08-1622">Tipo</span><span class="sxs-lookup"><span data-stu-id="28b08-1622">Type</span></span>|<span data-ttu-id="28b08-1623">Atributos</span><span class="sxs-lookup"><span data-stu-id="28b08-1623">Attributes</span></span>|<span data-ttu-id="28b08-1624">Descrição</span><span class="sxs-lookup"><span data-stu-id="28b08-1624">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="28b08-1625">String</span><span class="sxs-lookup"><span data-stu-id="28b08-1625">String</span></span>||<span data-ttu-id="28b08-p202">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="28b08-p202">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="28b08-1629">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1629">Object</span></span>|<span data-ttu-id="28b08-1630">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1630">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1631">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="28b08-1631">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="28b08-1632">Objeto</span><span class="sxs-lookup"><span data-stu-id="28b08-1632">Object</span></span>|<span data-ttu-id="28b08-1633">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1633">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1634">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="28b08-1634">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="28b08-1635">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="28b08-1635">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="28b08-1636">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="28b08-1636">&lt;optional&gt;</span></span>|<span data-ttu-id="28b08-1637">Se `text`, o estilo atual é aplicado nos clientes do Outlook na Web e do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="28b08-1637">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="28b08-1638">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="28b08-1638">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="28b08-1639">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook na Web e o estilo padrão é aplicado nos clientes do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="28b08-1639">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="28b08-1640">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="28b08-1640">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="28b08-1641">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="28b08-1641">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="28b08-1642">function</span><span class="sxs-lookup"><span data-stu-id="28b08-1642">function</span></span>||<span data-ttu-id="28b08-1643">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="28b08-1643">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28b08-1644">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28b08-1644">Requirements</span></span>

|<span data-ttu-id="28b08-1645">Requisito</span><span class="sxs-lookup"><span data-stu-id="28b08-1645">Requirement</span></span>|<span data-ttu-id="28b08-1646">Valor</span><span class="sxs-lookup"><span data-stu-id="28b08-1646">Value</span></span>|
|---|---|
|[<span data-ttu-id="28b08-1647">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28b08-1647">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="28b08-1648">1.2</span><span class="sxs-lookup"><span data-stu-id="28b08-1648">1.2</span></span>|
|[<span data-ttu-id="28b08-1649">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="28b08-1649">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="28b08-1650">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="28b08-1650">ReadWriteItem</span></span>|
|[<span data-ttu-id="28b08-1651">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28b08-1651">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="28b08-1652">Escrever</span><span class="sxs-lookup"><span data-stu-id="28b08-1652">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="28b08-1653">Exemplo</span><span class="sxs-lookup"><span data-stu-id="28b08-1653">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
