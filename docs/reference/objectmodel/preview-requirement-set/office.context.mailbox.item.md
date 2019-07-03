---
title: Office. Context. Mailbox. Item-visualização do conjunto de requisitos
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: 537ac59649b149d9bb54b09f8e16704adb813f58
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454899"
---
# <a name="item"></a><span data-ttu-id="45d29-102">item</span><span class="sxs-lookup"><span data-stu-id="45d29-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="45d29-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="45d29-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="45d29-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="45d29-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-106">Requirements</span></span>

|<span data-ttu-id="45d29-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-107">Requirement</span></span>|<span data-ttu-id="45d29-108">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-110">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-110">1.0</span></span>|
|[<span data-ttu-id="45d29-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="45d29-112">Restricted</span></span>|
|[<span data-ttu-id="45d29-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="45d29-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="45d29-115">Members and methods</span></span>

| <span data-ttu-id="45d29-116">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-116">Member</span></span> | <span data-ttu-id="45d29-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="45d29-118">attachments</span><span class="sxs-lookup"><span data-stu-id="45d29-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="45d29-119">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-119">Member</span></span> |
| [<span data-ttu-id="45d29-120">bcc</span><span class="sxs-lookup"><span data-stu-id="45d29-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="45d29-121">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-121">Member</span></span> |
| [<span data-ttu-id="45d29-122">body</span><span class="sxs-lookup"><span data-stu-id="45d29-122">body</span></span>](#body-body) | <span data-ttu-id="45d29-123">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-123">Member</span></span> |
| [<span data-ttu-id="45d29-124">Categorias</span><span class="sxs-lookup"><span data-stu-id="45d29-124">categories</span></span>](#categories-categories) | <span data-ttu-id="45d29-125">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-125">Member</span></span> |
| [<span data-ttu-id="45d29-126">cc</span><span class="sxs-lookup"><span data-stu-id="45d29-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="45d29-127">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-127">Member</span></span> |
| [<span data-ttu-id="45d29-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="45d29-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="45d29-129">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-129">Member</span></span> |
| [<span data-ttu-id="45d29-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="45d29-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="45d29-131">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-131">Member</span></span> |
| [<span data-ttu-id="45d29-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="45d29-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="45d29-133">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-133">Member</span></span> |
| [<span data-ttu-id="45d29-134">end</span><span class="sxs-lookup"><span data-stu-id="45d29-134">end</span></span>](#end-datetime) | <span data-ttu-id="45d29-135">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-135">Member</span></span> |
| [<span data-ttu-id="45d29-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="45d29-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="45d29-137">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-137">Member</span></span> |
| [<span data-ttu-id="45d29-138">from</span><span class="sxs-lookup"><span data-stu-id="45d29-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="45d29-139">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-139">Member</span></span> |
| [<span data-ttu-id="45d29-140">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="45d29-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="45d29-141">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-141">Member</span></span> |
| [<span data-ttu-id="45d29-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="45d29-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="45d29-143">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-143">Member</span></span> |
| [<span data-ttu-id="45d29-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="45d29-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="45d29-145">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-145">Member</span></span> |
| [<span data-ttu-id="45d29-146">itemId</span><span class="sxs-lookup"><span data-stu-id="45d29-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="45d29-147">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-147">Member</span></span> |
| [<span data-ttu-id="45d29-148">itemType</span><span class="sxs-lookup"><span data-stu-id="45d29-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="45d29-149">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-149">Member</span></span> |
| [<span data-ttu-id="45d29-150">location</span><span class="sxs-lookup"><span data-stu-id="45d29-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="45d29-151">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-151">Member</span></span> |
| [<span data-ttu-id="45d29-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="45d29-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="45d29-153">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-153">Member</span></span> |
| [<span data-ttu-id="45d29-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="45d29-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="45d29-155">Member</span><span class="sxs-lookup"><span data-stu-id="45d29-155">Member</span></span> |
| [<span data-ttu-id="45d29-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="45d29-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="45d29-157">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-157">Member</span></span> |
| [<span data-ttu-id="45d29-158">organizer</span><span class="sxs-lookup"><span data-stu-id="45d29-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="45d29-159">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-159">Member</span></span> |
| [<span data-ttu-id="45d29-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="45d29-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="45d29-161">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-161">Member</span></span> |
| [<span data-ttu-id="45d29-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="45d29-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="45d29-163">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-163">Member</span></span> |
| [<span data-ttu-id="45d29-164">sender</span><span class="sxs-lookup"><span data-stu-id="45d29-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="45d29-165">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-165">Member</span></span> |
| [<span data-ttu-id="45d29-166">seriesid</span><span class="sxs-lookup"><span data-stu-id="45d29-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="45d29-167">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-167">Member</span></span> |
| [<span data-ttu-id="45d29-168">start</span><span class="sxs-lookup"><span data-stu-id="45d29-168">start</span></span>](#start-datetime) | <span data-ttu-id="45d29-169">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-169">Member</span></span> |
| [<span data-ttu-id="45d29-170">subject</span><span class="sxs-lookup"><span data-stu-id="45d29-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="45d29-171">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-171">Member</span></span> |
| [<span data-ttu-id="45d29-172">to</span><span class="sxs-lookup"><span data-stu-id="45d29-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="45d29-173">Membro</span><span class="sxs-lookup"><span data-stu-id="45d29-173">Member</span></span> |
| [<span data-ttu-id="45d29-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="45d29-175">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-175">Method</span></span> |
| [<span data-ttu-id="45d29-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="45d29-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="45d29-177">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-177">Method</span></span> |
| [<span data-ttu-id="45d29-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="45d29-179">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-179">Method</span></span> |
| [<span data-ttu-id="45d29-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="45d29-181">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-181">Method</span></span> |
| [<span data-ttu-id="45d29-182">close</span><span class="sxs-lookup"><span data-stu-id="45d29-182">close</span></span>](#close) | <span data-ttu-id="45d29-183">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-183">Method</span></span> |
| [<span data-ttu-id="45d29-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="45d29-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="45d29-185">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-185">Method</span></span> |
| [<span data-ttu-id="45d29-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="45d29-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="45d29-187">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-187">Method</span></span> |
| [<span data-ttu-id="45d29-188">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="45d29-189">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-189">Method</span></span> |
| [<span data-ttu-id="45d29-190">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="45d29-191">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-191">Method</span></span> |
| [<span data-ttu-id="45d29-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="45d29-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="45d29-193">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-193">Method</span></span> |
| [<span data-ttu-id="45d29-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="45d29-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="45d29-195">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-195">Method</span></span> |
| [<span data-ttu-id="45d29-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="45d29-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="45d29-197">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-197">Method</span></span> |
| [<span data-ttu-id="45d29-198">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="45d29-199">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-199">Method</span></span> |
| [<span data-ttu-id="45d29-200">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="45d29-201">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-201">Method</span></span> |
| [<span data-ttu-id="45d29-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="45d29-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="45d29-203">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-203">Method</span></span> |
| [<span data-ttu-id="45d29-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="45d29-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="45d29-205">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-205">Method</span></span> |
| [<span data-ttu-id="45d29-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="45d29-207">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-207">Method</span></span> |
| [<span data-ttu-id="45d29-208">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="45d29-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="45d29-209">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-209">Method</span></span> |
| [<span data-ttu-id="45d29-210">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="45d29-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="45d29-211">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-211">Method</span></span> |
| [<span data-ttu-id="45d29-212">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="45d29-213">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-213">Method</span></span> |
| [<span data-ttu-id="45d29-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="45d29-215">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-215">Method</span></span> |
| [<span data-ttu-id="45d29-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="45d29-217">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-217">Method</span></span> |
| [<span data-ttu-id="45d29-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="45d29-219">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-219">Method</span></span> |
| [<span data-ttu-id="45d29-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="45d29-221">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-221">Method</span></span> |
| [<span data-ttu-id="45d29-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="45d29-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="45d29-223">Método</span><span class="sxs-lookup"><span data-stu-id="45d29-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="45d29-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-224">Example</span></span>

<span data-ttu-id="45d29-225">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="45d29-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="45d29-226">Membros</span><span class="sxs-lookup"><span data-stu-id="45d29-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="45d29-227">anexos: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="45d29-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="45d29-228">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="45d29-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="45d29-229">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="45d29-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-230">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="45d29-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="45d29-231">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="45d29-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-232">Type</span></span>

*   <span data-ttu-id="45d29-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="45d29-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-234">Requirements</span></span>

|<span data-ttu-id="45d29-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-235">Requirement</span></span>|<span data-ttu-id="45d29-236">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-238">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-238">1.0</span></span>|
|[<span data-ttu-id="45d29-239">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-240">ReadItem</span></span>|
|[<span data-ttu-id="45d29-241">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-242">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-243">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-243">Example</span></span>

<span data-ttu-id="45d29-244">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="45d29-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="45d29-245">CCO: [destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45d29-245">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="45d29-246">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="45d29-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="45d29-247">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="45d29-247">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-248">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-248">Type</span></span>

*   [<span data-ttu-id="45d29-249">Destinatários</span><span class="sxs-lookup"><span data-stu-id="45d29-249">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="45d29-250">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-250">Requirements</span></span>

|<span data-ttu-id="45d29-251">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-251">Requirement</span></span>|<span data-ttu-id="45d29-252">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-253">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-254">1.1</span><span class="sxs-lookup"><span data-stu-id="45d29-254">1.1</span></span>|
|[<span data-ttu-id="45d29-255">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-256">ReadItem</span></span>|
|[<span data-ttu-id="45d29-257">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-258">Escrever</span><span class="sxs-lookup"><span data-stu-id="45d29-258">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-259">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-259">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

---
---

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="45d29-260">corpo: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="45d29-260">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="45d29-261">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="45d29-261">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-262">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-262">Type</span></span>

*   [<span data-ttu-id="45d29-263">Body</span><span class="sxs-lookup"><span data-stu-id="45d29-263">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="45d29-264">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-264">Requirements</span></span>

|<span data-ttu-id="45d29-265">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-265">Requirement</span></span>|<span data-ttu-id="45d29-266">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-267">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-268">1.1</span><span class="sxs-lookup"><span data-stu-id="45d29-268">1.1</span></span>|
|[<span data-ttu-id="45d29-269">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-270">ReadItem</span></span>|
|[<span data-ttu-id="45d29-271">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-272">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-273">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-273">Example</span></span>

<span data-ttu-id="45d29-274">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="45d29-274">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="45d29-275">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-275">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="45d29-276">Categorias: [categorias](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="45d29-276">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="45d29-277">Obtém um objeto que fornece métodos para gerenciar as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="45d29-277">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-278">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="45d29-278">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-279">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-279">Type</span></span>

*   [<span data-ttu-id="45d29-280">Categories</span><span class="sxs-lookup"><span data-stu-id="45d29-280">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="45d29-281">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-281">Requirements</span></span>

|<span data-ttu-id="45d29-282">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-282">Requirement</span></span>|<span data-ttu-id="45d29-283">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-284">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-285">Visualização</span><span class="sxs-lookup"><span data-stu-id="45d29-285">Preview</span></span>|
|[<span data-ttu-id="45d29-286">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-287">ReadItem</span></span>|
|[<span data-ttu-id="45d29-288">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-289">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-289">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-290">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-290">Example</span></span>

<span data-ttu-id="45d29-291">Este exemplo obtém as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="45d29-291">This example gets the item's categories.</span></span>

```javascript
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="45d29-292">[destinatários](/javascript/api/outlook/office.recipients) [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|CC: Array. <</span><span class="sxs-lookup"><span data-stu-id="45d29-292">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="45d29-293">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="45d29-293">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="45d29-294">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="45d29-294">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45d29-295">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="45d29-295">Read mode</span></span>

<span data-ttu-id="45d29-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="45d29-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="45d29-298">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="45d29-298">Compose mode</span></span>

<span data-ttu-id="45d29-299">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="45d29-299">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="45d29-300">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-300">Type</span></span>

*   <span data-ttu-id="45d29-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45d29-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-302">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-302">Requirements</span></span>

|<span data-ttu-id="45d29-303">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-303">Requirement</span></span>|<span data-ttu-id="45d29-304">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-305">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-306">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-306">1.0</span></span>|
|[<span data-ttu-id="45d29-307">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-308">ReadItem</span></span>|
|[<span data-ttu-id="45d29-309">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-310">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-310">Compose or Read</span></span>|

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="45d29-311">(Nullable) Conversation: String</span><span class="sxs-lookup"><span data-stu-id="45d29-311">(nullable) conversationId: String</span></span>

<span data-ttu-id="45d29-312">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="45d29-312">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="45d29-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="45d29-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="45d29-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="45d29-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-317">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-317">Type</span></span>

*   <span data-ttu-id="45d29-318">String</span><span class="sxs-lookup"><span data-stu-id="45d29-318">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-319">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-319">Requirements</span></span>

|<span data-ttu-id="45d29-320">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-320">Requirement</span></span>|<span data-ttu-id="45d29-321">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-321">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-322">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-322">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-323">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-323">1.0</span></span>|
|[<span data-ttu-id="45d29-324">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-324">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-325">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-325">ReadItem</span></span>|
|[<span data-ttu-id="45d29-326">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-326">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-327">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-327">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-328">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-328">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="45d29-329">dateTimeCreated: data</span><span class="sxs-lookup"><span data-stu-id="45d29-329">dateTimeCreated: Date</span></span>

<span data-ttu-id="45d29-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="45d29-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-332">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-332">Type</span></span>

*   <span data-ttu-id="45d29-333">Data</span><span class="sxs-lookup"><span data-stu-id="45d29-333">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-334">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-334">Requirements</span></span>

|<span data-ttu-id="45d29-335">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-335">Requirement</span></span>|<span data-ttu-id="45d29-336">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-337">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-338">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-338">1.0</span></span>|
|[<span data-ttu-id="45d29-339">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-340">ReadItem</span></span>|
|[<span data-ttu-id="45d29-341">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-342">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-343">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-343">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="45d29-344">dateTimeModified: data</span><span class="sxs-lookup"><span data-stu-id="45d29-344">dateTimeModified: Date</span></span>

<span data-ttu-id="45d29-345">Obtém a data e a hora em que um item foi alterado pela última vez.</span><span class="sxs-lookup"><span data-stu-id="45d29-345">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="45d29-346">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="45d29-346">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-347">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="45d29-347">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-348">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-348">Type</span></span>

*   <span data-ttu-id="45d29-349">Data</span><span class="sxs-lookup"><span data-stu-id="45d29-349">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-350">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-350">Requirements</span></span>

|<span data-ttu-id="45d29-351">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-351">Requirement</span></span>|<span data-ttu-id="45d29-352">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-353">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-354">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-354">1.0</span></span>|
|[<span data-ttu-id="45d29-355">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-356">ReadItem</span></span>|
|[<span data-ttu-id="45d29-357">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-358">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-358">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-359">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-359">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="45d29-360">fim: data | [Tempo](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="45d29-360">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="45d29-361">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="45d29-361">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="45d29-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="45d29-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45d29-364">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="45d29-364">Read mode</span></span>

<span data-ttu-id="45d29-365">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="45d29-365">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="45d29-366">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="45d29-366">Compose mode</span></span>

<span data-ttu-id="45d29-367">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="45d29-367">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="45d29-368">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="45d29-368">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="45d29-369">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="45d29-369">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="45d29-370">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-370">Type</span></span>

*   <span data-ttu-id="45d29-371">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="45d29-371">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-372">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-372">Requirements</span></span>

|<span data-ttu-id="45d29-373">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-373">Requirement</span></span>|<span data-ttu-id="45d29-374">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-375">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-376">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-376">1.0</span></span>|
|[<span data-ttu-id="45d29-377">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-378">ReadItem</span></span>|
|[<span data-ttu-id="45d29-379">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-380">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-380">Compose or Read</span></span>|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="45d29-381">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="45d29-381">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="45d29-382">Obtém ou define os locais de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-382">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45d29-383">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="45d29-383">Read mode</span></span>

<span data-ttu-id="45d29-384">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que permite que você obtenha o conjunto de locais (cada um representado por um objeto [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associado ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="45d29-385">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="45d29-385">Compose mode</span></span>

<span data-ttu-id="45d29-386">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que fornece métodos para obter, remover ou adicionar locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-386">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-387">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-387">Type</span></span>

*   [<span data-ttu-id="45d29-388">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="45d29-388">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="45d29-389">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-389">Requirements</span></span>

|<span data-ttu-id="45d29-390">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-390">Requirement</span></span>|<span data-ttu-id="45d29-391">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-391">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-392">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-392">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-393">Visualização</span><span class="sxs-lookup"><span data-stu-id="45d29-393">Preview</span></span>|
|[<span data-ttu-id="45d29-394">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-395">ReadItem</span></span>|
|[<span data-ttu-id="45d29-396">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-396">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-397">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-397">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-398">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-398">Example</span></span>

<span data-ttu-id="45d29-399">O exemplo a seguir obtém os locais atuais associados ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-399">The following example gets the current locations associated with the appointment.</span></span>

```javascript
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

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="45d29-400">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="45d29-400">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="45d29-401">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="45d29-401">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="45d29-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="45d29-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-404">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="45d29-404">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45d29-405">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="45d29-405">Read mode</span></span>

<span data-ttu-id="45d29-406">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="45d29-406">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="45d29-407">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="45d29-407">Compose mode</span></span>

<span data-ttu-id="45d29-408">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="45d29-408">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="45d29-409">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-409">Type</span></span>

*   <span data-ttu-id="45d29-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="45d29-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-411">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-411">Requirements</span></span>

|<span data-ttu-id="45d29-412">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-412">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="45d29-413">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-414">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-414">1.0</span></span>|<span data-ttu-id="45d29-415">1.7</span><span class="sxs-lookup"><span data-stu-id="45d29-415">1.7</span></span>|
|[<span data-ttu-id="45d29-416">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-417">ReadItem</span></span>|<span data-ttu-id="45d29-418">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45d29-418">ReadWriteItem</span></span>|
|[<span data-ttu-id="45d29-419">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-419">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-420">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-420">Read</span></span>|<span data-ttu-id="45d29-421">Escrever</span><span class="sxs-lookup"><span data-stu-id="45d29-421">Compose</span></span>|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="45d29-422">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="45d29-422">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="45d29-423">Obtém ou define cabeçalhos de Internet personalizados em uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="45d29-423">Gets or sets custom internet headers on a message.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-424">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-424">Type</span></span>

*   [<span data-ttu-id="45d29-425">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="45d29-425">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="45d29-426">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-426">Requirements</span></span>

|<span data-ttu-id="45d29-427">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-427">Requirement</span></span>|<span data-ttu-id="45d29-428">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-429">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-430">Visualização</span><span class="sxs-lookup"><span data-stu-id="45d29-430">Preview</span></span>|
|[<span data-ttu-id="45d29-431">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-432">ReadItem</span></span>|
|[<span data-ttu-id="45d29-433">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-434">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-435">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-435">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="45d29-436">internetMessageId: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="45d29-436">internetMessageId: String</span></span>

<span data-ttu-id="45d29-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="45d29-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-439">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-439">Type</span></span>

*   <span data-ttu-id="45d29-440">String</span><span class="sxs-lookup"><span data-stu-id="45d29-440">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-441">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-441">Requirements</span></span>

|<span data-ttu-id="45d29-442">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-442">Requirement</span></span>|<span data-ttu-id="45d29-443">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-444">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-445">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-445">1.0</span></span>|
|[<span data-ttu-id="45d29-446">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-447">ReadItem</span></span>|
|[<span data-ttu-id="45d29-448">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-449">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-449">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-450">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-450">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="45d29-451">doclass: String</span><span class="sxs-lookup"><span data-stu-id="45d29-451">itemClass: String</span></span>

<span data-ttu-id="45d29-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="45d29-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="45d29-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="45d29-456">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-456">Type</span></span>|<span data-ttu-id="45d29-457">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-457">Description</span></span>|<span data-ttu-id="45d29-458">classe de item</span><span class="sxs-lookup"><span data-stu-id="45d29-458">item class</span></span>|
|---|---|---|
|<span data-ttu-id="45d29-459">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="45d29-459">Appointment items</span></span>|<span data-ttu-id="45d29-460">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="45d29-460">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="45d29-461">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="45d29-461">Message items</span></span>|<span data-ttu-id="45d29-462">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="45d29-462">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="45d29-463">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="45d29-463">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-464">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-464">Type</span></span>

*   <span data-ttu-id="45d29-465">String</span><span class="sxs-lookup"><span data-stu-id="45d29-465">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-466">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-466">Requirements</span></span>

|<span data-ttu-id="45d29-467">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-467">Requirement</span></span>|<span data-ttu-id="45d29-468">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-469">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-470">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-470">1.0</span></span>|
|[<span data-ttu-id="45d29-471">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-472">ReadItem</span></span>|
|[<span data-ttu-id="45d29-473">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-474">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-474">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-475">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-475">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="45d29-476">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="45d29-476">(nullable) itemId: String</span></span>

<span data-ttu-id="45d29-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="45d29-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-479">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="45d29-479">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="45d29-480">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="45d29-480">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="45d29-481">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="45d29-481">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="45d29-482">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="45d29-482">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="45d29-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-485">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-485">Type</span></span>

*   <span data-ttu-id="45d29-486">String</span><span class="sxs-lookup"><span data-stu-id="45d29-486">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-487">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-487">Requirements</span></span>

|<span data-ttu-id="45d29-488">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-488">Requirement</span></span>|<span data-ttu-id="45d29-489">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-490">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-490">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-491">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-491">1.0</span></span>|
|[<span data-ttu-id="45d29-492">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-492">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-493">ReadItem</span></span>|
|[<span data-ttu-id="45d29-494">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-494">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-495">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-495">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-496">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-496">Example</span></span>

<span data-ttu-id="45d29-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="45d29-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="45d29-499">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="45d29-499">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="45d29-500">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="45d29-500">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="45d29-501">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-501">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-502">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-502">Type</span></span>

*   [<span data-ttu-id="45d29-503">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="45d29-503">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="45d29-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-504">Requirements</span></span>

|<span data-ttu-id="45d29-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-505">Requirement</span></span>|<span data-ttu-id="45d29-506">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-507">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-508">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-508">1.0</span></span>|
|[<span data-ttu-id="45d29-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-510">ReadItem</span></span>|
|[<span data-ttu-id="45d29-511">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-512">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-512">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-513">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-513">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="45d29-514">local: cadeia de caracteres | [Local](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="45d29-514">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="45d29-515">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-515">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45d29-516">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="45d29-516">Read mode</span></span>

<span data-ttu-id="45d29-517">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-517">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="45d29-518">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="45d29-518">Compose mode</span></span>

<span data-ttu-id="45d29-519">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-519">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="45d29-520">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-520">Type</span></span>

*   <span data-ttu-id="45d29-521">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="45d29-521">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-522">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-522">Requirements</span></span>

|<span data-ttu-id="45d29-523">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-523">Requirement</span></span>|<span data-ttu-id="45d29-524">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-525">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-526">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-526">1.0</span></span>|
|[<span data-ttu-id="45d29-527">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-527">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-528">ReadItem</span></span>|
|[<span data-ttu-id="45d29-529">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-530">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-530">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="45d29-531">normalizedSubject: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="45d29-531">normalizedSubject: String</span></span>

<span data-ttu-id="45d29-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="45d29-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="45d29-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="45d29-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-536">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-536">Type</span></span>

*   <span data-ttu-id="45d29-537">String</span><span class="sxs-lookup"><span data-stu-id="45d29-537">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-538">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-538">Requirements</span></span>

|<span data-ttu-id="45d29-539">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-539">Requirement</span></span>|<span data-ttu-id="45d29-540">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-541">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-542">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-542">1.0</span></span>|
|[<span data-ttu-id="45d29-543">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-544">ReadItem</span></span>|
|[<span data-ttu-id="45d29-545">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-546">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-547">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-547">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="45d29-548">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="45d29-548">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="45d29-549">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="45d29-549">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-550">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-550">Type</span></span>

*   [<span data-ttu-id="45d29-551">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="45d29-551">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="45d29-552">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-552">Requirements</span></span>

|<span data-ttu-id="45d29-553">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-553">Requirement</span></span>|<span data-ttu-id="45d29-554">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-554">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-555">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-555">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-556">1.3</span><span class="sxs-lookup"><span data-stu-id="45d29-556">1.3</span></span>|
|[<span data-ttu-id="45d29-557">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-557">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-558">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-558">ReadItem</span></span>|
|[<span data-ttu-id="45d29-559">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-559">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-560">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-560">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-561">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-561">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="45d29-562">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[](/javascript/api/outlook/office.recipients) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="45d29-562">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="45d29-563">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="45d29-563">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="45d29-564">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="45d29-564">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45d29-565">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="45d29-565">Read mode</span></span>

<span data-ttu-id="45d29-566">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="45d29-566">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="45d29-567">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="45d29-567">Compose mode</span></span>

<span data-ttu-id="45d29-568">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="45d29-568">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="45d29-569">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-569">Type</span></span>

*   <span data-ttu-id="45d29-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45d29-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-571">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-571">Requirements</span></span>

|<span data-ttu-id="45d29-572">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-572">Requirement</span></span>|<span data-ttu-id="45d29-573">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-573">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-574">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-575">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-575">1.0</span></span>|
|[<span data-ttu-id="45d29-576">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-577">ReadItem</span></span>|
|[<span data-ttu-id="45d29-578">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-579">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-579">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="45d29-580">organizador: [](/javascript/api/outlook/office.emailaddressdetails)|[organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="45d29-580">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="45d29-581">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="45d29-581">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45d29-582">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="45d29-582">Read mode</span></span>

<span data-ttu-id="45d29-583">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="45d29-583">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="45d29-584">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="45d29-584">Compose mode</span></span>

<span data-ttu-id="45d29-585">A `organizer` propriedade retorna um [](/javascript/api/outlook/office.organizer) objeto organizador que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="45d29-585">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="45d29-586">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-586">Type</span></span>

*   <span data-ttu-id="45d29-587">[](/javascript/api/outlook/office.emailaddressdetails) | [Organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="45d29-587">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-588">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-588">Requirements</span></span>

|<span data-ttu-id="45d29-589">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-589">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="45d29-590">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-591">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-591">1.0</span></span>|<span data-ttu-id="45d29-592">1.7</span><span class="sxs-lookup"><span data-stu-id="45d29-592">1.7</span></span>|
|[<span data-ttu-id="45d29-593">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-594">ReadItem</span></span>|<span data-ttu-id="45d29-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45d29-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="45d29-596">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-597">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-597">Read</span></span>|<span data-ttu-id="45d29-598">Escrever</span><span class="sxs-lookup"><span data-stu-id="45d29-598">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="45d29-599">(anulável) recorrência [](/javascript/api/outlook/office.recurrence) : recorrência</span><span class="sxs-lookup"><span data-stu-id="45d29-599">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="45d29-600">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-600">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="45d29-601">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="45d29-601">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="45d29-602">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-602">Read and compose modes for appointment items.</span></span> <span data-ttu-id="45d29-603">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="45d29-603">Read mode for meeting request items.</span></span>

<span data-ttu-id="45d29-604">A `recurrence` propriedade retorna um [](/javascript/api/outlook/office.recurrence) objeto de recorrência para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="45d29-604">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="45d29-605">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="45d29-605">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="45d29-606">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="45d29-606">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="45d29-607">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="45d29-607">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="45d29-608">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="45d29-608">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45d29-609">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="45d29-609">Read mode</span></span>

<span data-ttu-id="45d29-610">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-610">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="45d29-611">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="45d29-611">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="45d29-612">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="45d29-612">Compose mode</span></span>

<span data-ttu-id="45d29-613">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-613">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="45d29-614">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="45d29-614">This is available for appointments.</span></span>

```javascript
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

##### <a name="type"></a><span data-ttu-id="45d29-615">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-615">Type</span></span>

* [<span data-ttu-id="45d29-616">Recorrência</span><span class="sxs-lookup"><span data-stu-id="45d29-616">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="45d29-617">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-617">Requirement</span></span>|<span data-ttu-id="45d29-618">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-619">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-620">1.7</span><span class="sxs-lookup"><span data-stu-id="45d29-620">1.7</span></span>|
|[<span data-ttu-id="45d29-621">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-622">ReadItem</span></span>|
|[<span data-ttu-id="45d29-623">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-624">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-624">Compose or Read</span></span>|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="45d29-625">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[](/javascript/api/outlook/office.recipients) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="45d29-625">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="45d29-626">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="45d29-626">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="45d29-627">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="45d29-627">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45d29-628">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="45d29-628">Read mode</span></span>

<span data-ttu-id="45d29-629">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="45d29-629">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="45d29-630">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="45d29-630">Compose mode</span></span>

<span data-ttu-id="45d29-631">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="45d29-631">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="45d29-632">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-632">Type</span></span>

*   <span data-ttu-id="45d29-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45d29-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-634">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-634">Requirements</span></span>

|<span data-ttu-id="45d29-635">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-635">Requirement</span></span>|<span data-ttu-id="45d29-636">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-637">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-638">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-638">1.0</span></span>|
|[<span data-ttu-id="45d29-639">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-640">ReadItem</span></span>|
|[<span data-ttu-id="45d29-641">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-642">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-642">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="45d29-643">remetente: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="45d29-643">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="45d29-p128">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="45d29-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="45d29-p129">As propriedades [`from`](#from-emailaddressdetailsfrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="45d29-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-648">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="45d29-648">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-649">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-649">Type</span></span>

*   [<span data-ttu-id="45d29-650">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="45d29-650">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="45d29-651">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-651">Requirements</span></span>

|<span data-ttu-id="45d29-652">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-652">Requirement</span></span>|<span data-ttu-id="45d29-653">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-654">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-655">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-655">1.0</span></span>|
|[<span data-ttu-id="45d29-656">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-657">ReadItem</span></span>|
|[<span data-ttu-id="45d29-658">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-659">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-659">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-660">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-660">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="45d29-661">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="45d29-661">(nullable) seriesId: String</span></span>

<span data-ttu-id="45d29-662">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="45d29-662">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="45d29-663">No Outlook na Web e clientes de desktop, o `seriesId` retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="45d29-663">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="45d29-664">No entanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="45d29-664">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-665">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="45d29-665">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="45d29-666">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="45d29-666">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="45d29-667">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="45d29-667">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="45d29-668">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="45d29-668">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="45d29-669">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="45d29-669">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="45d29-670">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-670">Type</span></span>

* <span data-ttu-id="45d29-671">String</span><span class="sxs-lookup"><span data-stu-id="45d29-671">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-672">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-672">Requirements</span></span>

|<span data-ttu-id="45d29-673">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-673">Requirement</span></span>|<span data-ttu-id="45d29-674">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-674">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-675">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-675">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-676">1.7</span><span class="sxs-lookup"><span data-stu-id="45d29-676">1.7</span></span>|
|[<span data-ttu-id="45d29-677">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-677">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-678">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-678">ReadItem</span></span>|
|[<span data-ttu-id="45d29-679">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-679">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-680">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-680">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-681">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-681">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

---
---

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="45d29-682">Início: data | [Tempo](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="45d29-682">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="45d29-683">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="45d29-683">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="45d29-p132">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="45d29-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45d29-686">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="45d29-686">Read mode</span></span>

<span data-ttu-id="45d29-687">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="45d29-687">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="45d29-688">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="45d29-688">Compose mode</span></span>

<span data-ttu-id="45d29-689">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="45d29-689">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="45d29-690">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="45d29-690">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="45d29-691">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="45d29-691">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="45d29-692">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-692">Type</span></span>

*   <span data-ttu-id="45d29-693">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="45d29-693">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-694">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-694">Requirements</span></span>

|<span data-ttu-id="45d29-695">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-695">Requirement</span></span>|<span data-ttu-id="45d29-696">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-697">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-698">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-698">1.0</span></span>|
|[<span data-ttu-id="45d29-699">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-700">ReadItem</span></span>|
|[<span data-ttu-id="45d29-701">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-702">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-702">Compose or Read</span></span>|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="45d29-703">subject: cadeia de caracteres | [Assunto](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="45d29-703">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="45d29-704">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="45d29-704">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="45d29-705">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="45d29-705">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45d29-706">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="45d29-706">Read mode</span></span>

<span data-ttu-id="45d29-p133">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="45d29-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="45d29-709">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="45d29-709">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="45d29-710">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="45d29-710">Compose mode</span></span>
<span data-ttu-id="45d29-711">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="45d29-711">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="45d29-712">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-712">Type</span></span>

*   <span data-ttu-id="45d29-713">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="45d29-713">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-714">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-714">Requirements</span></span>

|<span data-ttu-id="45d29-715">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-715">Requirement</span></span>|<span data-ttu-id="45d29-716">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-716">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-717">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-717">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-718">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-718">1.0</span></span>|
|[<span data-ttu-id="45d29-719">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-719">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-720">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-720">ReadItem</span></span>|
|[<span data-ttu-id="45d29-721">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-721">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-722">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-722">Compose or Read</span></span>|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="45d29-723">para: Array. <[](/javascript/api/outlook/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook/office.recipients) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="45d29-723">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="45d29-724">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="45d29-724">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="45d29-725">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="45d29-725">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45d29-726">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="45d29-726">Read mode</span></span>

<span data-ttu-id="45d29-p135">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="45d29-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="45d29-729">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="45d29-729">Compose mode</span></span>

<span data-ttu-id="45d29-730">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="45d29-730">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="45d29-731">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-731">Type</span></span>

*   <span data-ttu-id="45d29-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45d29-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-733">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-733">Requirements</span></span>

|<span data-ttu-id="45d29-734">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-734">Requirement</span></span>|<span data-ttu-id="45d29-735">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-735">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-736">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-736">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-737">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-737">1.0</span></span>|
|[<span data-ttu-id="45d29-738">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-738">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-739">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-739">ReadItem</span></span>|
|[<span data-ttu-id="45d29-740">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-740">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-741">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-741">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="45d29-742">Métodos</span><span class="sxs-lookup"><span data-stu-id="45d29-742">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="45d29-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="45d29-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="45d29-744">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="45d29-744">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="45d29-745">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="45d29-745">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="45d29-746">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="45d29-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-747">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-747">Parameters</span></span>
|<span data-ttu-id="45d29-748">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-748">Name</span></span>|<span data-ttu-id="45d29-749">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-749">Type</span></span>|<span data-ttu-id="45d29-750">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-750">Attributes</span></span>|<span data-ttu-id="45d29-751">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-751">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="45d29-752">String</span><span class="sxs-lookup"><span data-stu-id="45d29-752">String</span></span>||<span data-ttu-id="45d29-p136">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="45d29-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="45d29-755">String</span><span class="sxs-lookup"><span data-stu-id="45d29-755">String</span></span>||<span data-ttu-id="45d29-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="45d29-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="45d29-758">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-758">Object</span></span>|<span data-ttu-id="45d29-759">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-759">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-760">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45d29-761">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-761">Object</span></span>|<span data-ttu-id="45d29-762">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-762">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-763">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="45d29-764">Booliano</span><span class="sxs-lookup"><span data-stu-id="45d29-764">Boolean</span></span>|<span data-ttu-id="45d29-765">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-765">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-766">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="45d29-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="45d29-767">function</span><span class="sxs-lookup"><span data-stu-id="45d29-767">function</span></span>|<span data-ttu-id="45d29-768">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-768">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-769">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="45d29-770">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45d29-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="45d29-771">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="45d29-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="45d29-772">Erros</span><span class="sxs-lookup"><span data-stu-id="45d29-772">Errors</span></span>

|<span data-ttu-id="45d29-773">Código de erro</span><span class="sxs-lookup"><span data-stu-id="45d29-773">Error code</span></span>|<span data-ttu-id="45d29-774">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="45d29-775">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="45d29-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="45d29-776">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="45d29-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="45d29-777">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="45d29-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-778">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-778">Requirements</span></span>

|<span data-ttu-id="45d29-779">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-779">Requirement</span></span>|<span data-ttu-id="45d29-780">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-781">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-782">1.1</span><span class="sxs-lookup"><span data-stu-id="45d29-782">1.1</span></span>|
|[<span data-ttu-id="45d29-783">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45d29-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="45d29-785">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-786">Escrever</span><span class="sxs-lookup"><span data-stu-id="45d29-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="45d29-787">Exemplos</span><span class="sxs-lookup"><span data-stu-id="45d29-787">Examples</span></span>

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

<span data-ttu-id="45d29-788">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="45d29-788">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="45d29-789">addFileAttachmentFromBase64Async (base64file, AttachmentName, [Options], [callback])</span><span class="sxs-lookup"><span data-stu-id="45d29-789">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="45d29-790">Adiciona um arquivo da codificação Base64 a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="45d29-790">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="45d29-791">O `addFileAttachmentFromBase64Async` método carrega o arquivo da codificação Base64 e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="45d29-791">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="45d29-792">Esse método retorna o identificador de anexo no objeto AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="45d29-792">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="45d29-793">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="45d29-793">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-794">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-794">Parameters</span></span>

|<span data-ttu-id="45d29-795">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-795">Name</span></span>|<span data-ttu-id="45d29-796">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-796">Type</span></span>|<span data-ttu-id="45d29-797">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-797">Attributes</span></span>|<span data-ttu-id="45d29-798">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-798">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="45d29-799">String</span><span class="sxs-lookup"><span data-stu-id="45d29-799">String</span></span>||<span data-ttu-id="45d29-800">O conteúdo codificado em Base64 de uma imagem ou arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="45d29-800">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="45d29-801">String</span><span class="sxs-lookup"><span data-stu-id="45d29-801">String</span></span>||<span data-ttu-id="45d29-p139">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="45d29-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="45d29-804">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-804">Object</span></span>|<span data-ttu-id="45d29-805">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-805">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-806">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-806">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45d29-807">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-807">Object</span></span>|<span data-ttu-id="45d29-808">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-808">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-809">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-809">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="45d29-810">Booliano</span><span class="sxs-lookup"><span data-stu-id="45d29-810">Boolean</span></span>|<span data-ttu-id="45d29-811">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-811">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-812">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="45d29-812">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="45d29-813">function</span><span class="sxs-lookup"><span data-stu-id="45d29-813">function</span></span>|<span data-ttu-id="45d29-814">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-814">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-815">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-815">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="45d29-816">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45d29-816">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="45d29-817">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="45d29-817">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="45d29-818">Erros</span><span class="sxs-lookup"><span data-stu-id="45d29-818">Errors</span></span>

|<span data-ttu-id="45d29-819">Código de erro</span><span class="sxs-lookup"><span data-stu-id="45d29-819">Error code</span></span>|<span data-ttu-id="45d29-820">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-820">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="45d29-821">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="45d29-821">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="45d29-822">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="45d29-822">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="45d29-823">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="45d29-823">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-824">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-824">Requirements</span></span>

|<span data-ttu-id="45d29-825">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-825">Requirement</span></span>|<span data-ttu-id="45d29-826">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-827">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-828">Visualização</span><span class="sxs-lookup"><span data-stu-id="45d29-828">Preview</span></span>|
|[<span data-ttu-id="45d29-829">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-830">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45d29-830">ReadWriteItem</span></span>|
|[<span data-ttu-id="45d29-831">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-832">Escrever</span><span class="sxs-lookup"><span data-stu-id="45d29-832">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="45d29-833">Exemplos</span><span class="sxs-lookup"><span data-stu-id="45d29-833">Examples</span></span>

```javascript
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

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="45d29-834">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="45d29-834">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="45d29-835">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="45d29-835">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="45d29-836">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="45d29-836">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-837">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-837">Parameters</span></span>

| <span data-ttu-id="45d29-838">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-838">Name</span></span> | <span data-ttu-id="45d29-839">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-839">Type</span></span> | <span data-ttu-id="45d29-840">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-840">Attributes</span></span> | <span data-ttu-id="45d29-841">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-841">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="45d29-842">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="45d29-842">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="45d29-843">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="45d29-843">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="45d29-844">Função</span><span class="sxs-lookup"><span data-stu-id="45d29-844">Function</span></span> || <span data-ttu-id="45d29-p140">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="45d29-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="45d29-848">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-848">Object</span></span> | <span data-ttu-id="45d29-849">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-849">&lt;optional&gt;</span></span> | <span data-ttu-id="45d29-850">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-850">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="45d29-851">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-851">Object</span></span> | <span data-ttu-id="45d29-852">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-852">&lt;optional&gt;</span></span> | <span data-ttu-id="45d29-853">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-853">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="45d29-854">function</span><span class="sxs-lookup"><span data-stu-id="45d29-854">function</span></span>| <span data-ttu-id="45d29-855">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-855">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-856">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-857">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-857">Requirements</span></span>

|<span data-ttu-id="45d29-858">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-858">Requirement</span></span>| <span data-ttu-id="45d29-859">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-859">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-860">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45d29-861">1.7</span><span class="sxs-lookup"><span data-stu-id="45d29-861">1.7</span></span> |
|[<span data-ttu-id="45d29-862">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45d29-863">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-863">ReadItem</span></span> |
|[<span data-ttu-id="45d29-864">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="45d29-865">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-865">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="45d29-866">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-866">Example</span></span>

```javascript
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="45d29-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="45d29-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="45d29-868">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-868">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="45d29-p141">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="45d29-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="45d29-872">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="45d29-872">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="45d29-873">Se o suplemento do Office estiver em execução no Outlook na Web, o `addItemAttachmentAsync` método poderá anexar itens a itens diferentes do item que você está editando; no entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="45d29-873">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-874">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-874">Parameters</span></span>

|<span data-ttu-id="45d29-875">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-875">Name</span></span>|<span data-ttu-id="45d29-876">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-876">Type</span></span>|<span data-ttu-id="45d29-877">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-877">Attributes</span></span>|<span data-ttu-id="45d29-878">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-878">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="45d29-879">String</span><span class="sxs-lookup"><span data-stu-id="45d29-879">String</span></span>||<span data-ttu-id="45d29-p142">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="45d29-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="45d29-882">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="45d29-882">String</span></span>||<span data-ttu-id="45d29-883">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="45d29-883">The subject of the item to be attached.</span></span> <span data-ttu-id="45d29-884">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="45d29-884">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="45d29-885">Object</span><span class="sxs-lookup"><span data-stu-id="45d29-885">Object</span></span>|<span data-ttu-id="45d29-886">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-886">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-887">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-887">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45d29-888">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-888">Object</span></span>|<span data-ttu-id="45d29-889">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-889">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-890">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-890">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45d29-891">function</span><span class="sxs-lookup"><span data-stu-id="45d29-891">function</span></span>|<span data-ttu-id="45d29-892">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-892">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-893">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-893">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="45d29-894">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45d29-894">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="45d29-895">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="45d29-895">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="45d29-896">Erros</span><span class="sxs-lookup"><span data-stu-id="45d29-896">Errors</span></span>

|<span data-ttu-id="45d29-897">Código de erro</span><span class="sxs-lookup"><span data-stu-id="45d29-897">Error code</span></span>|<span data-ttu-id="45d29-898">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-898">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="45d29-899">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="45d29-899">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-900">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-900">Requirements</span></span>

|<span data-ttu-id="45d29-901">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-901">Requirement</span></span>|<span data-ttu-id="45d29-902">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-903">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-904">1.1</span><span class="sxs-lookup"><span data-stu-id="45d29-904">1.1</span></span>|
|[<span data-ttu-id="45d29-905">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45d29-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="45d29-907">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-908">Escrever</span><span class="sxs-lookup"><span data-stu-id="45d29-908">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-909">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-909">Example</span></span>

<span data-ttu-id="45d29-910">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="45d29-910">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

---
---

#### <a name="close"></a><span data-ttu-id="45d29-911">close()</span><span class="sxs-lookup"><span data-stu-id="45d29-911">close()</span></span>

<span data-ttu-id="45d29-912">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="45d29-912">Closes the current item that is being composed.</span></span>

<span data-ttu-id="45d29-p144">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="45d29-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-915">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="45d29-915">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="45d29-916">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="45d29-916">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-917">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-917">Requirements</span></span>

|<span data-ttu-id="45d29-918">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-918">Requirement</span></span>|<span data-ttu-id="45d29-919">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-919">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-920">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-920">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-921">1.3</span><span class="sxs-lookup"><span data-stu-id="45d29-921">1.3</span></span>|
|[<span data-ttu-id="45d29-922">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-922">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-923">Restrito</span><span class="sxs-lookup"><span data-stu-id="45d29-923">Restricted</span></span>|
|[<span data-ttu-id="45d29-924">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-924">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-925">Escrever</span><span class="sxs-lookup"><span data-stu-id="45d29-925">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="45d29-926">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="45d29-926">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="45d29-927">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="45d29-927">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-928">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="45d29-928">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="45d29-929">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="45d29-929">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="45d29-930">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="45d29-930">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="45d29-931">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="45d29-931">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="45d29-932">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="45d29-932">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="45d29-933">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="45d29-933">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-934">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-934">Parameters</span></span>

|<span data-ttu-id="45d29-935">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-935">Name</span></span>|<span data-ttu-id="45d29-936">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-936">Type</span></span>|<span data-ttu-id="45d29-937">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-937">Attributes</span></span>|<span data-ttu-id="45d29-938">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-938">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="45d29-939">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="45d29-939">String &#124; Object</span></span>||<span data-ttu-id="45d29-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="45d29-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="45d29-942">**OU**</span><span class="sxs-lookup"><span data-stu-id="45d29-942">**OR**</span></span><br/><span data-ttu-id="45d29-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="45d29-945">String</span><span class="sxs-lookup"><span data-stu-id="45d29-945">String</span></span>|<span data-ttu-id="45d29-946">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-946">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="45d29-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="45d29-949">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-949">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="45d29-950">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-950">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-951">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="45d29-951">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="45d29-952">String</span><span class="sxs-lookup"><span data-stu-id="45d29-952">String</span></span>||<span data-ttu-id="45d29-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="45d29-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="45d29-955">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="45d29-955">String</span></span>||<span data-ttu-id="45d29-956">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="45d29-956">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="45d29-957">String</span><span class="sxs-lookup"><span data-stu-id="45d29-957">String</span></span>||<span data-ttu-id="45d29-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="45d29-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="45d29-960">Booliano</span><span class="sxs-lookup"><span data-stu-id="45d29-960">Boolean</span></span>||<span data-ttu-id="45d29-p151">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="45d29-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="45d29-963">String</span><span class="sxs-lookup"><span data-stu-id="45d29-963">String</span></span>||<span data-ttu-id="45d29-p152">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="45d29-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="45d29-967">function</span><span class="sxs-lookup"><span data-stu-id="45d29-967">function</span></span>|<span data-ttu-id="45d29-968">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-968">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-969">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-969">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-970">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-970">Requirements</span></span>

|<span data-ttu-id="45d29-971">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-971">Requirement</span></span>|<span data-ttu-id="45d29-972">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-972">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-973">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-973">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-974">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-974">1.0</span></span>|
|[<span data-ttu-id="45d29-975">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-975">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-976">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-976">ReadItem</span></span>|
|[<span data-ttu-id="45d29-977">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-977">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-978">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-978">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="45d29-979">Exemplos</span><span class="sxs-lookup"><span data-stu-id="45d29-979">Examples</span></span>

<span data-ttu-id="45d29-980">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="45d29-980">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="45d29-981">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="45d29-981">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="45d29-982">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="45d29-982">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="45d29-983">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="45d29-983">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="45d29-984">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="45d29-984">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="45d29-985">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-985">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="45d29-986">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="45d29-986">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="45d29-987">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="45d29-987">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-988">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="45d29-988">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="45d29-989">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="45d29-989">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="45d29-990">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="45d29-990">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="45d29-991">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="45d29-991">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="45d29-992">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="45d29-992">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="45d29-993">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="45d29-993">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-994">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-994">Parameters</span></span>

|<span data-ttu-id="45d29-995">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-995">Name</span></span>|<span data-ttu-id="45d29-996">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-996">Type</span></span>|<span data-ttu-id="45d29-997">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-997">Attributes</span></span>|<span data-ttu-id="45d29-998">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-998">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="45d29-999">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="45d29-999">String &#124; Object</span></span>||<span data-ttu-id="45d29-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="45d29-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="45d29-1002">**OU**</span><span class="sxs-lookup"><span data-stu-id="45d29-1002">**OR**</span></span><br/><span data-ttu-id="45d29-p155">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="45d29-1005">String</span><span class="sxs-lookup"><span data-stu-id="45d29-1005">String</span></span>|<span data-ttu-id="45d29-1006">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-p156">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="45d29-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="45d29-1009">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1009">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="45d29-1010">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1011">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="45d29-1011">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="45d29-1012">String</span><span class="sxs-lookup"><span data-stu-id="45d29-1012">String</span></span>||<span data-ttu-id="45d29-p157">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="45d29-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="45d29-1015">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="45d29-1015">String</span></span>||<span data-ttu-id="45d29-1016">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="45d29-1016">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="45d29-1017">String</span><span class="sxs-lookup"><span data-stu-id="45d29-1017">String</span></span>||<span data-ttu-id="45d29-p158">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="45d29-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="45d29-1020">Booliano</span><span class="sxs-lookup"><span data-stu-id="45d29-1020">Boolean</span></span>||<span data-ttu-id="45d29-p159">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="45d29-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="45d29-1023">String</span><span class="sxs-lookup"><span data-stu-id="45d29-1023">String</span></span>||<span data-ttu-id="45d29-p160">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="45d29-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="45d29-1027">function</span><span class="sxs-lookup"><span data-stu-id="45d29-1027">function</span></span>|<span data-ttu-id="45d29-1028">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1028">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1029">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-1029">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1030">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1030">Requirements</span></span>

|<span data-ttu-id="45d29-1031">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1031">Requirement</span></span>|<span data-ttu-id="45d29-1032">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1032">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1033">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1033">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1034">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-1034">1.0</span></span>|
|[<span data-ttu-id="45d29-1035">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1035">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1036">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1036">ReadItem</span></span>|
|[<span data-ttu-id="45d29-1037">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-1037">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1038">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-1038">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="45d29-1039">Exemplos</span><span class="sxs-lookup"><span data-stu-id="45d29-1039">Examples</span></span>

<span data-ttu-id="45d29-1040">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="45d29-1040">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="45d29-1041">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="45d29-1041">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="45d29-1042">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="45d29-1042">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="45d29-1043">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="45d29-1043">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="45d29-1044">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="45d29-1044">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="45d29-1045">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1045">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="45d29-1046">getAttachmentContentAsync (attachmentid, [opções], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="45d29-1046">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="45d29-1047">Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um `AttachmentContent` objeto.</span><span class="sxs-lookup"><span data-stu-id="45d29-1047">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="45d29-1048">O `getAttachmentContentAsync` método obtém o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="45d29-1048">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="45d29-1049">Como prática recomendada, você deve usar o identificador para recuperar um anexo na mesma sessão em que o attachmentIds foi recuperado com a `getAttachmentsAsync` chamada ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="45d29-1049">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="45d29-1050">No Outlook na Web e dispositivos móveis, o identificador de anexo é válido somente dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="45d29-1050">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="45d29-1051">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1051">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1052">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1052">Parameters</span></span>

|<span data-ttu-id="45d29-1053">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1053">Name</span></span>|<span data-ttu-id="45d29-1054">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1054">Type</span></span>|<span data-ttu-id="45d29-1055">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-1055">Attributes</span></span>|<span data-ttu-id="45d29-1056">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1056">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="45d29-1057">String</span><span class="sxs-lookup"><span data-stu-id="45d29-1057">String</span></span>||<span data-ttu-id="45d29-1058">O identificador do anexo que você deseja obter.</span><span class="sxs-lookup"><span data-stu-id="45d29-1058">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="45d29-1059">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1059">Object</span></span>|<span data-ttu-id="45d29-1060">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1060">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1061">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-1061">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45d29-1062">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1062">Object</span></span>|<span data-ttu-id="45d29-1063">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1064">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1064">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45d29-1065">function</span><span class="sxs-lookup"><span data-stu-id="45d29-1065">function</span></span>|<span data-ttu-id="45d29-1066">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1067">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-1067">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1068">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1068">Requirements</span></span>

|<span data-ttu-id="45d29-1069">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1069">Requirement</span></span>|<span data-ttu-id="45d29-1070">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1070">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1071">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1071">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1072">Visualização</span><span class="sxs-lookup"><span data-stu-id="45d29-1072">Preview</span></span>|
|[<span data-ttu-id="45d29-1073">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1073">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1074">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1074">ReadItem</span></span>|
|[<span data-ttu-id="45d29-1075">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-1075">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1076">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-1076">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45d29-1077">Retorna:</span><span class="sxs-lookup"><span data-stu-id="45d29-1077">Returns:</span></span>

<span data-ttu-id="45d29-1078">Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="45d29-1078">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="45d29-1079">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1079">Example</span></span>

```javascript
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
  if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    // Handle file attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
    // Handle email item attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
    // Handle .icalender attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    // Handle cloud attachment.
  } else {
    // Handle attachment formats that are not supported.
  }
}
```

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="45d29-1080">getAttachmentsAsync ([Options], [callback]) → array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="45d29-1080">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="45d29-1081">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="45d29-1081">Gets the item's attachments as an array.</span></span> <span data-ttu-id="45d29-1082">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="45d29-1082">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1083">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1083">Parameters</span></span>

|<span data-ttu-id="45d29-1084">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1084">Name</span></span>|<span data-ttu-id="45d29-1085">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1085">Type</span></span>|<span data-ttu-id="45d29-1086">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-1086">Attributes</span></span>|<span data-ttu-id="45d29-1087">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1087">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="45d29-1088">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1088">Object</span></span>|<span data-ttu-id="45d29-1089">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1089">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1090">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-1090">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45d29-1091">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1091">Object</span></span>|<span data-ttu-id="45d29-1092">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1093">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1093">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45d29-1094">function</span><span class="sxs-lookup"><span data-stu-id="45d29-1094">function</span></span>|<span data-ttu-id="45d29-1095">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1096">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-1096">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1097">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1097">Requirements</span></span>

|<span data-ttu-id="45d29-1098">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1098">Requirement</span></span>|<span data-ttu-id="45d29-1099">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1099">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1100">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1100">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1101">Visualização</span><span class="sxs-lookup"><span data-stu-id="45d29-1101">Preview</span></span>|
|[<span data-ttu-id="45d29-1102">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1102">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1103">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1103">ReadItem</span></span>|
|[<span data-ttu-id="45d29-1104">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-1104">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1105">Escrever</span><span class="sxs-lookup"><span data-stu-id="45d29-1105">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="45d29-1106">Retorna:</span><span class="sxs-lookup"><span data-stu-id="45d29-1106">Returns:</span></span>

<span data-ttu-id="45d29-1107">Tipo: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="45d29-1107">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="45d29-1108">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1108">Example</span></span>

<span data-ttu-id="45d29-1109">O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="45d29-1109">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="45d29-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="45d29-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="45d29-1111">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="45d29-1111">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-1112">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="45d29-1112">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-1113">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1113">Requirements</span></span>

|<span data-ttu-id="45d29-1114">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1114">Requirement</span></span>|<span data-ttu-id="45d29-1115">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1115">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1116">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1116">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1117">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-1117">1.0</span></span>|
|[<span data-ttu-id="45d29-1118">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1118">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1119">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1119">ReadItem</span></span>|
|[<span data-ttu-id="45d29-1120">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-1120">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1121">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-1121">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45d29-1122">Retorna:</span><span class="sxs-lookup"><span data-stu-id="45d29-1122">Returns:</span></span>

<span data-ttu-id="45d29-1123">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="45d29-1123">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="45d29-1124">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1124">Example</span></span>

<span data-ttu-id="45d29-1125">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="45d29-1125">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="45d29-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="45d29-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="45d29-1127">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="45d29-1127">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-1128">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="45d29-1128">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1129">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1129">Parameters</span></span>

|<span data-ttu-id="45d29-1130">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1130">Name</span></span>|<span data-ttu-id="45d29-1131">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1131">Type</span></span>|<span data-ttu-id="45d29-1132">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1132">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="45d29-1133">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="45d29-1133">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="45d29-1134">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="45d29-1134">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1135">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1135">Requirements</span></span>

|<span data-ttu-id="45d29-1136">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1136">Requirement</span></span>|<span data-ttu-id="45d29-1137">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1138">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1139">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-1139">1.0</span></span>|
|[<span data-ttu-id="45d29-1140">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1141">Restrito</span><span class="sxs-lookup"><span data-stu-id="45d29-1141">Restricted</span></span>|
|[<span data-ttu-id="45d29-1142">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1143">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45d29-1144">Retorna:</span><span class="sxs-lookup"><span data-stu-id="45d29-1144">Returns:</span></span>

<span data-ttu-id="45d29-1145">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="45d29-1145">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="45d29-1146">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="45d29-1146">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="45d29-1147">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="45d29-1147">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="45d29-1148">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-1148">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="45d29-1149">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="45d29-1149">Value of `entityType`</span></span>|<span data-ttu-id="45d29-1150">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="45d29-1150">Type of objects in returned array</span></span>|<span data-ttu-id="45d29-1151">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="45d29-1151">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="45d29-1152">String</span><span class="sxs-lookup"><span data-stu-id="45d29-1152">String</span></span>|<span data-ttu-id="45d29-1153">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="45d29-1153">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="45d29-1154">Contato</span><span class="sxs-lookup"><span data-stu-id="45d29-1154">Contact</span></span>|<span data-ttu-id="45d29-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="45d29-1155">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="45d29-1156">String</span><span class="sxs-lookup"><span data-stu-id="45d29-1156">String</span></span>|<span data-ttu-id="45d29-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="45d29-1157">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="45d29-1158">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="45d29-1158">MeetingSuggestion</span></span>|<span data-ttu-id="45d29-1159">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="45d29-1159">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="45d29-1160">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="45d29-1160">PhoneNumber</span></span>|<span data-ttu-id="45d29-1161">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="45d29-1161">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="45d29-1162">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="45d29-1162">TaskSuggestion</span></span>|<span data-ttu-id="45d29-1163">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="45d29-1163">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="45d29-1164">String</span><span class="sxs-lookup"><span data-stu-id="45d29-1164">String</span></span>|<span data-ttu-id="45d29-1165">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="45d29-1165">**Restricted**</span></span>|

<span data-ttu-id="45d29-1166">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="45d29-1166">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="45d29-1167">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1167">Example</span></span>

<span data-ttu-id="45d29-1168">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="45d29-1168">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="45d29-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="45d29-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="45d29-1170">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="45d29-1170">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-1171">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="45d29-1171">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="45d29-1172">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="45d29-1172">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1173">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1173">Parameters</span></span>

|<span data-ttu-id="45d29-1174">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1174">Name</span></span>|<span data-ttu-id="45d29-1175">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1175">Type</span></span>|<span data-ttu-id="45d29-1176">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1176">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="45d29-1177">String</span><span class="sxs-lookup"><span data-stu-id="45d29-1177">String</span></span>|<span data-ttu-id="45d29-1178">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="45d29-1178">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1179">Requirements</span></span>

|<span data-ttu-id="45d29-1180">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1180">Requirement</span></span>|<span data-ttu-id="45d29-1181">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1181">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1183">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-1183">1.0</span></span>|
|[<span data-ttu-id="45d29-1184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1185">ReadItem</span></span>|
|[<span data-ttu-id="45d29-1186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-1186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1187">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-1187">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45d29-1188">Retorna:</span><span class="sxs-lookup"><span data-stu-id="45d29-1188">Returns:</span></span>

<span data-ttu-id="45d29-p164">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="45d29-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="45d29-1191">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="45d29-1191">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="45d29-1192">getInitializationContextAsync ([opções], [callback])</span><span class="sxs-lookup"><span data-stu-id="45d29-1192">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="45d29-1193">Obtém dados de inicialização passados quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="45d29-1193">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-1194">Este método só é compatível com o Outlook 2016 ou posterior no Windows (clique para executar versões posteriores a 16.0.8413.1000) e Outlook na Web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="45d29-1194">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1195">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1195">Parameters</span></span>

|<span data-ttu-id="45d29-1196">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1196">Name</span></span>|<span data-ttu-id="45d29-1197">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1197">Type</span></span>|<span data-ttu-id="45d29-1198">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-1198">Attributes</span></span>|<span data-ttu-id="45d29-1199">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1199">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="45d29-1200">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1200">Object</span></span>|<span data-ttu-id="45d29-1201">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1201">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1202">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-1202">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45d29-1203">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1203">Object</span></span>|<span data-ttu-id="45d29-1204">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1204">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1205">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1205">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45d29-1206">function</span><span class="sxs-lookup"><span data-stu-id="45d29-1206">function</span></span>|<span data-ttu-id="45d29-1207">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1207">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1208">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-1208">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="45d29-1209">Com êxito, os dados de inicialização são fornecidos na `asyncResult.value` Propriedade como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="45d29-1209">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="45d29-1210">Se não houver nenhum contexto de inicialização, `asyncResult` o objeto conterá `Error` um objeto com `code` sua propriedade definida `9020` como e `name` sua propriedade definida `GenericResponseError`como.</span><span class="sxs-lookup"><span data-stu-id="45d29-1210">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1211">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1211">Requirements</span></span>

|<span data-ttu-id="45d29-1212">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1212">Requirement</span></span>|<span data-ttu-id="45d29-1213">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1213">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1214">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1215">Visualização</span><span class="sxs-lookup"><span data-stu-id="45d29-1215">Preview</span></span>|
|[<span data-ttu-id="45d29-1216">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1217">ReadItem</span></span>|
|[<span data-ttu-id="45d29-1218">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-1218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1219">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-1219">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-1220">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1220">Example</span></span>

```javascript
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

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="45d29-1221">getItemIdAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="45d29-1221">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="45d29-1222">Obtém de forma assíncrona a ID de um item salvo.</span><span class="sxs-lookup"><span data-stu-id="45d29-1222">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="45d29-1223">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="45d29-1223">Compose mode only.</span></span>

<span data-ttu-id="45d29-1224">Quando invocado, este método retorna a ID do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1224">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-1225">Se seu suplemento chamar `getItemIdAsync` um item no modo de redação (por exemplo, para `itemId` usar com o EWS ou a API REST), lembre-se de que, quando o Outlook estiver no modo cache, pode levar algum tempo para que o item seja sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="45d29-1225">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="45d29-1226">Até que o item seja sincronizado, `itemId` o não é reconhecido e usado retorna um erro.</span><span class="sxs-lookup"><span data-stu-id="45d29-1226">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1227">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1227">Parameters</span></span>

|<span data-ttu-id="45d29-1228">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1228">Name</span></span>|<span data-ttu-id="45d29-1229">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1229">Type</span></span>|<span data-ttu-id="45d29-1230">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-1230">Attributes</span></span>|<span data-ttu-id="45d29-1231">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1231">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="45d29-1232">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1232">Object</span></span>|<span data-ttu-id="45d29-1233">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1233">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1234">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-1234">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45d29-1235">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1235">Object</span></span>|<span data-ttu-id="45d29-1236">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1236">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1237">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1237">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45d29-1238">function</span><span class="sxs-lookup"><span data-stu-id="45d29-1238">function</span></span>||<span data-ttu-id="45d29-1239">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-1239">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="45d29-1240">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45d29-1240">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="45d29-1241">Erros</span><span class="sxs-lookup"><span data-stu-id="45d29-1241">Errors</span></span>

|<span data-ttu-id="45d29-1242">Código de erro</span><span class="sxs-lookup"><span data-stu-id="45d29-1242">Error code</span></span>|<span data-ttu-id="45d29-1243">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1243">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="45d29-1244">A ID não pode ser recuperada até que o item seja salvo.</span><span class="sxs-lookup"><span data-stu-id="45d29-1244">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1245">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1245">Requirements</span></span>

|<span data-ttu-id="45d29-1246">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1246">Requirement</span></span>|<span data-ttu-id="45d29-1247">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1247">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1248">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1248">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1249">Visualização</span><span class="sxs-lookup"><span data-stu-id="45d29-1249">Preview</span></span>|
|[<span data-ttu-id="45d29-1250">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1250">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1251">ReadItem</span></span>|
|[<span data-ttu-id="45d29-1252">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-1252">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1253">Escrever</span><span class="sxs-lookup"><span data-stu-id="45d29-1253">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="45d29-1254">Exemplos</span><span class="sxs-lookup"><span data-stu-id="45d29-1254">Examples</span></span>

```javascript
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="45d29-1255">O exemplo a seguir mostra a estrutura do `result` parâmetro que é passado para a função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1255">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="45d29-1256">A `value` propriedade contém a ID do item.</span><span class="sxs-lookup"><span data-stu-id="45d29-1256">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="45d29-1257">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="45d29-1257">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="45d29-1258">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="45d29-1258">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-1259">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="45d29-1259">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="45d29-p168">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="45d29-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="45d29-1263">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="45d29-1263">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="45d29-1264">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="45d29-1264">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="45d29-p169">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="45d29-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-1268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1268">Requirements</span></span>

|<span data-ttu-id="45d29-1269">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1269">Requirement</span></span>|<span data-ttu-id="45d29-1270">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1271">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-1272">1.0</span></span>|
|[<span data-ttu-id="45d29-1273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1274">ReadItem</span></span>|
|[<span data-ttu-id="45d29-1275">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-1275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1276">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45d29-1277">Retorna:</span><span class="sxs-lookup"><span data-stu-id="45d29-1277">Returns:</span></span>

<span data-ttu-id="45d29-p170">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="45d29-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="45d29-1280">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="45d29-1280">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="45d29-1281">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1281">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="45d29-1282">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1282">Example</span></span>

<span data-ttu-id="45d29-1283">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="45d29-1283">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="45d29-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="45d29-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="45d29-1285">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="45d29-1285">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-1286">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="45d29-1286">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="45d29-1287">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="45d29-1287">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="45d29-p171">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="45d29-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1290">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1290">Parameters</span></span>

|<span data-ttu-id="45d29-1291">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1291">Name</span></span>|<span data-ttu-id="45d29-1292">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1292">Type</span></span>|<span data-ttu-id="45d29-1293">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1293">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="45d29-1294">String</span><span class="sxs-lookup"><span data-stu-id="45d29-1294">String</span></span>|<span data-ttu-id="45d29-1295">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="45d29-1295">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1296">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1296">Requirements</span></span>

|<span data-ttu-id="45d29-1297">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1297">Requirement</span></span>|<span data-ttu-id="45d29-1298">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1298">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1299">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1299">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1300">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-1300">1.0</span></span>|
|[<span data-ttu-id="45d29-1301">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1301">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1302">ReadItem</span></span>|
|[<span data-ttu-id="45d29-1303">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-1303">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1304">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-1304">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45d29-1305">Retorna:</span><span class="sxs-lookup"><span data-stu-id="45d29-1305">Returns:</span></span>

<span data-ttu-id="45d29-1306">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="45d29-1306">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="45d29-1307">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="45d29-1307">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="45d29-1308">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="45d29-1308">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="45d29-1309">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1309">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="45d29-1310">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="45d29-1310">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="45d29-1311">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="45d29-1311">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="45d29-p172">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="45d29-p172">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1314">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1314">Parameters</span></span>

|<span data-ttu-id="45d29-1315">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1315">Name</span></span>|<span data-ttu-id="45d29-1316">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1316">Type</span></span>|<span data-ttu-id="45d29-1317">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-1317">Attributes</span></span>|<span data-ttu-id="45d29-1318">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1318">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="45d29-1319">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="45d29-1319">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="45d29-p173">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="45d29-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="45d29-1323">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1323">Object</span></span>|<span data-ttu-id="45d29-1324">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1324">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1325">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-1325">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45d29-1326">Object</span><span class="sxs-lookup"><span data-stu-id="45d29-1326">Object</span></span>|<span data-ttu-id="45d29-1327">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1327">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1328">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1328">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45d29-1329">function</span><span class="sxs-lookup"><span data-stu-id="45d29-1329">function</span></span>||<span data-ttu-id="45d29-1330">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-1330">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="45d29-1331">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="45d29-1331">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="45d29-1332">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="45d29-1332">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1333">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1333">Requirements</span></span>

|<span data-ttu-id="45d29-1334">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1334">Requirement</span></span>|<span data-ttu-id="45d29-1335">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1335">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1336">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1337">1.2</span><span class="sxs-lookup"><span data-stu-id="45d29-1337">1.2</span></span>|
|[<span data-ttu-id="45d29-1338">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1338">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1339">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1339">ReadWriteItem</span></span>|
|[<span data-ttu-id="45d29-1340">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-1340">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1341">Escrever</span><span class="sxs-lookup"><span data-stu-id="45d29-1341">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="45d29-1342">Retorna:</span><span class="sxs-lookup"><span data-stu-id="45d29-1342">Returns:</span></span>

<span data-ttu-id="45d29-1343">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="45d29-1343">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="45d29-1344">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="45d29-1344">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="45d29-1345">String</span><span class="sxs-lookup"><span data-stu-id="45d29-1345">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="45d29-1346">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1346">Example</span></span>

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

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="45d29-1347">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="45d29-1347">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="45d29-1348">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="45d29-1348">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="45d29-1349">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="45d29-1349">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-1350">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="45d29-1350">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-1351">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1351">Requirements</span></span>

|<span data-ttu-id="45d29-1352">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1352">Requirement</span></span>|<span data-ttu-id="45d29-1353">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1353">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1354">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1354">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1355">1.6</span><span class="sxs-lookup"><span data-stu-id="45d29-1355">1.6</span></span>|
|[<span data-ttu-id="45d29-1356">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1356">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1357">ReadItem</span></span>|
|[<span data-ttu-id="45d29-1358">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-1358">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1359">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-1359">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45d29-1360">Retorna:</span><span class="sxs-lookup"><span data-stu-id="45d29-1360">Returns:</span></span>

<span data-ttu-id="45d29-1361">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="45d29-1361">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="45d29-1362">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1362">Example</span></span>

<span data-ttu-id="45d29-1363">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="45d29-1363">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="45d29-1364">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="45d29-1364">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="45d29-p176">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="45d29-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-1367">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="45d29-1367">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="45d29-p177">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="45d29-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="45d29-1371">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="45d29-1371">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="45d29-1372">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="45d29-1372">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="45d29-p178">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="45d29-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="45d29-1376">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1376">Requirements</span></span>

|<span data-ttu-id="45d29-1377">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1377">Requirement</span></span>|<span data-ttu-id="45d29-1378">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1378">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1379">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1380">1.6</span><span class="sxs-lookup"><span data-stu-id="45d29-1380">1.6</span></span>|
|[<span data-ttu-id="45d29-1381">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1381">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1382">ReadItem</span></span>|
|[<span data-ttu-id="45d29-1383">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-1383">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1384">Read</span><span class="sxs-lookup"><span data-stu-id="45d29-1384">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45d29-1385">Retorna:</span><span class="sxs-lookup"><span data-stu-id="45d29-1385">Returns:</span></span>

<span data-ttu-id="45d29-p179">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="45d29-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="45d29-1388">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1388">Example</span></span>

<span data-ttu-id="45d29-1389">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="45d29-1389">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="45d29-1390">getSharedPropertiesAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="45d29-1390">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="45d29-1391">Obtém as propriedades do compromisso ou da mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="45d29-1391">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1392">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1392">Parameters</span></span>

|<span data-ttu-id="45d29-1393">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1393">Name</span></span>|<span data-ttu-id="45d29-1394">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1394">Type</span></span>|<span data-ttu-id="45d29-1395">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-1395">Attributes</span></span>|<span data-ttu-id="45d29-1396">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1396">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="45d29-1397">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1397">Object</span></span>|<span data-ttu-id="45d29-1398">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1398">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1399">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-1399">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45d29-1400">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1400">Object</span></span>|<span data-ttu-id="45d29-1401">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1401">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1402">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1402">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45d29-1403">function</span><span class="sxs-lookup"><span data-stu-id="45d29-1403">function</span></span>||<span data-ttu-id="45d29-1404">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-1404">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="45d29-1405">As propriedades compartilhadas são fornecidas [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) como um objeto `asyncResult.value` na propriedade.</span><span class="sxs-lookup"><span data-stu-id="45d29-1405">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="45d29-1406">Este objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="45d29-1406">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1407">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1407">Requirements</span></span>

|<span data-ttu-id="45d29-1408">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1408">Requirement</span></span>|<span data-ttu-id="45d29-1409">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1409">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1410">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1411">Visualização</span><span class="sxs-lookup"><span data-stu-id="45d29-1411">Preview</span></span>|
|[<span data-ttu-id="45d29-1412">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1413">ReadItem</span></span>|
|[<span data-ttu-id="45d29-1414">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-1414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1415">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-1415">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-1416">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1416">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="45d29-1417">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="45d29-1417">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="45d29-1418">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="45d29-1418">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="45d29-p181">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="45d29-p181">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1422">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1422">Parameters</span></span>

|<span data-ttu-id="45d29-1423">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1423">Name</span></span>|<span data-ttu-id="45d29-1424">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1424">Type</span></span>|<span data-ttu-id="45d29-1425">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-1425">Attributes</span></span>|<span data-ttu-id="45d29-1426">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1426">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="45d29-1427">function</span><span class="sxs-lookup"><span data-stu-id="45d29-1427">function</span></span>||<span data-ttu-id="45d29-1428">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-1428">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="45d29-1429">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45d29-1429">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="45d29-1430">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="45d29-1430">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="45d29-1431">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1431">Object</span></span>|<span data-ttu-id="45d29-1432">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1432">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1433">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1433">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="45d29-1434">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1434">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1435">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1435">Requirements</span></span>

|<span data-ttu-id="45d29-1436">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1436">Requirement</span></span>|<span data-ttu-id="45d29-1437">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1437">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1438">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1439">1.0</span><span class="sxs-lookup"><span data-stu-id="45d29-1439">1.0</span></span>|
|[<span data-ttu-id="45d29-1440">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1440">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1441">ReadItem</span></span>|
|[<span data-ttu-id="45d29-1442">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-1442">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1443">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-1443">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-1444">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1444">Example</span></span>

<span data-ttu-id="45d29-p184">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="45d29-p184">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="45d29-1448">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="45d29-1448">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="45d29-1449">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="45d29-1449">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="45d29-1450">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="45d29-1450">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="45d29-1451">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="45d29-1451">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="45d29-1452">No Outlook na Web e dispositivos móveis, o identificador de anexo é válido somente dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="45d29-1452">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="45d29-1453">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1453">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1454">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1454">Parameters</span></span>

|<span data-ttu-id="45d29-1455">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1455">Name</span></span>|<span data-ttu-id="45d29-1456">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1456">Type</span></span>|<span data-ttu-id="45d29-1457">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-1457">Attributes</span></span>|<span data-ttu-id="45d29-1458">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1458">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="45d29-1459">String</span><span class="sxs-lookup"><span data-stu-id="45d29-1459">String</span></span>||<span data-ttu-id="45d29-1460">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="45d29-1460">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="45d29-1461">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1461">Object</span></span>|<span data-ttu-id="45d29-1462">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1462">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1463">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-1463">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45d29-1464">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1464">Object</span></span>|<span data-ttu-id="45d29-1465">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1465">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1466">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1466">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45d29-1467">function</span><span class="sxs-lookup"><span data-stu-id="45d29-1467">function</span></span>|<span data-ttu-id="45d29-1468">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1468">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1469">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-1469">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="45d29-1470">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="45d29-1470">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="45d29-1471">Erros</span><span class="sxs-lookup"><span data-stu-id="45d29-1471">Errors</span></span>

|<span data-ttu-id="45d29-1472">Código de erro</span><span class="sxs-lookup"><span data-stu-id="45d29-1472">Error code</span></span>|<span data-ttu-id="45d29-1473">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1473">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="45d29-1474">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="45d29-1474">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1475">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1475">Requirements</span></span>

|<span data-ttu-id="45d29-1476">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1476">Requirement</span></span>|<span data-ttu-id="45d29-1477">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1477">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1478">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1478">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1479">1.1</span><span class="sxs-lookup"><span data-stu-id="45d29-1479">1.1</span></span>|
|[<span data-ttu-id="45d29-1480">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1480">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1481">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1481">ReadWriteItem</span></span>|
|[<span data-ttu-id="45d29-1482">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-1482">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1483">Escrever</span><span class="sxs-lookup"><span data-stu-id="45d29-1483">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-1484">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1484">Example</span></span>

<span data-ttu-id="45d29-1485">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="45d29-1485">The following code removes an attachment with an identifier of '0'.</span></span>

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

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="45d29-1486">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="45d29-1486">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="45d29-1487">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="45d29-1487">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="45d29-1488">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="45d29-1488">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1489">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1489">Parameters</span></span>

| <span data-ttu-id="45d29-1490">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1490">Name</span></span> | <span data-ttu-id="45d29-1491">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1491">Type</span></span> | <span data-ttu-id="45d29-1492">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-1492">Attributes</span></span> | <span data-ttu-id="45d29-1493">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1493">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="45d29-1494">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="45d29-1494">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="45d29-1495">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="45d29-1495">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="45d29-1496">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1496">Object</span></span> | <span data-ttu-id="45d29-1497">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1497">&lt;optional&gt;</span></span> | <span data-ttu-id="45d29-1498">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-1498">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="45d29-1499">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1499">Object</span></span> | <span data-ttu-id="45d29-1500">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1500">&lt;optional&gt;</span></span> | <span data-ttu-id="45d29-1501">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1501">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="45d29-1502">function</span><span class="sxs-lookup"><span data-stu-id="45d29-1502">function</span></span>| <span data-ttu-id="45d29-1503">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1503">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1504">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-1504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1505">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1505">Requirements</span></span>

|<span data-ttu-id="45d29-1506">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1506">Requirement</span></span>| <span data-ttu-id="45d29-1507">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1507">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1508">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45d29-1509">1.7</span><span class="sxs-lookup"><span data-stu-id="45d29-1509">1.7</span></span> |
|[<span data-ttu-id="45d29-1510">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45d29-1511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1511">ReadItem</span></span> |
|[<span data-ttu-id="45d29-1512">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="45d29-1512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="45d29-1513">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="45d29-1513">Compose or Read</span></span> |

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="45d29-1514">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="45d29-1514">saveAsync([options], callback)</span></span>

<span data-ttu-id="45d29-1515">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="45d29-1515">Asynchronously saves an item.</span></span>

<span data-ttu-id="45d29-1516">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1516">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="45d29-1517">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="45d29-1517">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="45d29-1518">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="45d29-1518">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-1519">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="45d29-1519">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="45d29-1520">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="45d29-1520">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="45d29-p188">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="45d29-p188">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="45d29-1524">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="45d29-1524">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="45d29-1525">O Outlook no Mac não dá suporte à gravação de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="45d29-1525">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="45d29-1526">O `saveAsync` método falha quando chamado de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="45d29-1526">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="45d29-1527">Consulte [não é possível salvar uma reunião como rascunho no Outlook para Mac usando a API do Office js](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="45d29-1527">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="45d29-1528">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="45d29-1528">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1529">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1529">Parameters</span></span>

|<span data-ttu-id="45d29-1530">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1530">Name</span></span>|<span data-ttu-id="45d29-1531">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1531">Type</span></span>|<span data-ttu-id="45d29-1532">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-1532">Attributes</span></span>|<span data-ttu-id="45d29-1533">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1533">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="45d29-1534">Object</span><span class="sxs-lookup"><span data-stu-id="45d29-1534">Object</span></span>|<span data-ttu-id="45d29-1535">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1535">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1536">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-1536">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45d29-1537">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1537">Object</span></span>|<span data-ttu-id="45d29-1538">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1538">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1539">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1539">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45d29-1540">function</span><span class="sxs-lookup"><span data-stu-id="45d29-1540">function</span></span>||<span data-ttu-id="45d29-1541">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-1541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="45d29-1542">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45d29-1542">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1543">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1543">Requirements</span></span>

|<span data-ttu-id="45d29-1544">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1544">Requirement</span></span>|<span data-ttu-id="45d29-1545">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1545">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1546">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1547">1.3</span><span class="sxs-lookup"><span data-stu-id="45d29-1547">1.3</span></span>|
|[<span data-ttu-id="45d29-1548">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1549">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1549">ReadWriteItem</span></span>|
|[<span data-ttu-id="45d29-1550">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-1550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1551">Escrever</span><span class="sxs-lookup"><span data-stu-id="45d29-1551">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="45d29-1552">Exemplos</span><span class="sxs-lookup"><span data-stu-id="45d29-1552">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="45d29-p190">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="45d29-p190">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="45d29-1555">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="45d29-1555">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="45d29-1556">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="45d29-1556">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="45d29-p191">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="45d29-p191">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45d29-1560">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="45d29-1560">Parameters</span></span>

|<span data-ttu-id="45d29-1561">Nome</span><span class="sxs-lookup"><span data-stu-id="45d29-1561">Name</span></span>|<span data-ttu-id="45d29-1562">Tipo</span><span class="sxs-lookup"><span data-stu-id="45d29-1562">Type</span></span>|<span data-ttu-id="45d29-1563">Atributos</span><span class="sxs-lookup"><span data-stu-id="45d29-1563">Attributes</span></span>|<span data-ttu-id="45d29-1564">Descrição</span><span class="sxs-lookup"><span data-stu-id="45d29-1564">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="45d29-1565">String</span><span class="sxs-lookup"><span data-stu-id="45d29-1565">String</span></span>||<span data-ttu-id="45d29-p192">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="45d29-p192">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="45d29-1569">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1569">Object</span></span>|<span data-ttu-id="45d29-1570">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1570">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1571">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="45d29-1571">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45d29-1572">Objeto</span><span class="sxs-lookup"><span data-stu-id="45d29-1572">Object</span></span>|<span data-ttu-id="45d29-1573">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1573">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1574">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="45d29-1574">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="45d29-1575">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="45d29-1575">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="45d29-1576">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="45d29-1576">&lt;optional&gt;</span></span>|<span data-ttu-id="45d29-1577">Se `text`, o estilo atual é aplicado no Outlook na Web e clientes da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="45d29-1577">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="45d29-1578">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="45d29-1578">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="45d29-1579">Se `html` e o campo oferecer suporte a HTML (o assunto não), o estilo atual será aplicado no Outlook na Web e o estilo padrão será aplicado nos clientes da área de trabalho do Outlook.</span><span class="sxs-lookup"><span data-stu-id="45d29-1579">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="45d29-1580">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="45d29-1580">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="45d29-1581">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="45d29-1581">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="45d29-1582">function</span><span class="sxs-lookup"><span data-stu-id="45d29-1582">function</span></span>||<span data-ttu-id="45d29-1583">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45d29-1583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45d29-1584">Requisitos</span><span class="sxs-lookup"><span data-stu-id="45d29-1584">Requirements</span></span>

|<span data-ttu-id="45d29-1585">Requisito</span><span class="sxs-lookup"><span data-stu-id="45d29-1585">Requirement</span></span>|<span data-ttu-id="45d29-1586">Valor</span><span class="sxs-lookup"><span data-stu-id="45d29-1586">Value</span></span>|
|---|---|
|[<span data-ttu-id="45d29-1587">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="45d29-1587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45d29-1588">1.2</span><span class="sxs-lookup"><span data-stu-id="45d29-1588">1.2</span></span>|
|[<span data-ttu-id="45d29-1589">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="45d29-1589">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45d29-1590">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45d29-1590">ReadWriteItem</span></span>|
|[<span data-ttu-id="45d29-1591">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="45d29-1591">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45d29-1592">Escrever</span><span class="sxs-lookup"><span data-stu-id="45d29-1592">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="45d29-1593">Exemplo</span><span class="sxs-lookup"><span data-stu-id="45d29-1593">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
