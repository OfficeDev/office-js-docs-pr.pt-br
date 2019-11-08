---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,8
description: ''
ms.date: 11/06/2019
localization_priority: Normal
ms.openlocfilehash: fe55299acc6fb10c6e0e6a4536c300c84a53664e
ms.sourcegitcommit: 08c0b9ff319c391922fa43d3c2e9783cf6b53b1b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/08/2019
ms.locfileid: "38066197"
---
# <a name="item"></a><span data-ttu-id="9a830-102">item</span><span class="sxs-lookup"><span data-stu-id="9a830-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="9a830-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="9a830-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="9a830-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="9a830-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-106">Requirements</span></span>

|<span data-ttu-id="9a830-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-107">Requirement</span></span>|<span data-ttu-id="9a830-108">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-110">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-110">1.0</span></span>|
|[<span data-ttu-id="9a830-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="9a830-112">Restricted</span></span>|
|[<span data-ttu-id="9a830-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9a830-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="9a830-115">Members and methods</span></span>

| <span data-ttu-id="9a830-116">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-116">Member</span></span> | <span data-ttu-id="9a830-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9a830-118">attachments</span><span class="sxs-lookup"><span data-stu-id="9a830-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="9a830-119">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-119">Member</span></span> |
| [<span data-ttu-id="9a830-120">bcc</span><span class="sxs-lookup"><span data-stu-id="9a830-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="9a830-121">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-121">Member</span></span> |
| [<span data-ttu-id="9a830-122">body</span><span class="sxs-lookup"><span data-stu-id="9a830-122">body</span></span>](#body-body) | <span data-ttu-id="9a830-123">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-123">Member</span></span> |
| [<span data-ttu-id="9a830-124">Categorias</span><span class="sxs-lookup"><span data-stu-id="9a830-124">categories</span></span>](#categories-categories) | <span data-ttu-id="9a830-125">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-125">Member</span></span> |
| [<span data-ttu-id="9a830-126">cc</span><span class="sxs-lookup"><span data-stu-id="9a830-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="9a830-127">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-127">Member</span></span> |
| [<span data-ttu-id="9a830-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="9a830-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="9a830-129">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-129">Member</span></span> |
| [<span data-ttu-id="9a830-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="9a830-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="9a830-131">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-131">Member</span></span> |
| [<span data-ttu-id="9a830-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="9a830-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="9a830-133">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-133">Member</span></span> |
| [<span data-ttu-id="9a830-134">end</span><span class="sxs-lookup"><span data-stu-id="9a830-134">end</span></span>](#end-datetime) | <span data-ttu-id="9a830-135">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-135">Member</span></span> |
| [<span data-ttu-id="9a830-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="9a830-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="9a830-137">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-137">Member</span></span> |
| [<span data-ttu-id="9a830-138">from</span><span class="sxs-lookup"><span data-stu-id="9a830-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="9a830-139">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-139">Member</span></span> |
| [<span data-ttu-id="9a830-140">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="9a830-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="9a830-141">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-141">Member</span></span> |
| [<span data-ttu-id="9a830-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="9a830-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="9a830-143">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-143">Member</span></span> |
| [<span data-ttu-id="9a830-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="9a830-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="9a830-145">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-145">Member</span></span> |
| [<span data-ttu-id="9a830-146">itemId</span><span class="sxs-lookup"><span data-stu-id="9a830-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="9a830-147">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-147">Member</span></span> |
| [<span data-ttu-id="9a830-148">itemType</span><span class="sxs-lookup"><span data-stu-id="9a830-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="9a830-149">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-149">Member</span></span> |
| [<span data-ttu-id="9a830-150">location</span><span class="sxs-lookup"><span data-stu-id="9a830-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="9a830-151">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-151">Member</span></span> |
| [<span data-ttu-id="9a830-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="9a830-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="9a830-153">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-153">Member</span></span> |
| [<span data-ttu-id="9a830-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="9a830-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="9a830-155">Member</span><span class="sxs-lookup"><span data-stu-id="9a830-155">Member</span></span> |
| [<span data-ttu-id="9a830-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="9a830-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="9a830-157">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-157">Member</span></span> |
| [<span data-ttu-id="9a830-158">organizer</span><span class="sxs-lookup"><span data-stu-id="9a830-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="9a830-159">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-159">Member</span></span> |
| [<span data-ttu-id="9a830-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="9a830-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="9a830-161">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-161">Member</span></span> |
| [<span data-ttu-id="9a830-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="9a830-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="9a830-163">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-163">Member</span></span> |
| [<span data-ttu-id="9a830-164">sender</span><span class="sxs-lookup"><span data-stu-id="9a830-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="9a830-165">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-165">Member</span></span> |
| [<span data-ttu-id="9a830-166">seriesid</span><span class="sxs-lookup"><span data-stu-id="9a830-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="9a830-167">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-167">Member</span></span> |
| [<span data-ttu-id="9a830-168">start</span><span class="sxs-lookup"><span data-stu-id="9a830-168">start</span></span>](#start-datetime) | <span data-ttu-id="9a830-169">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-169">Member</span></span> |
| [<span data-ttu-id="9a830-170">subject</span><span class="sxs-lookup"><span data-stu-id="9a830-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="9a830-171">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-171">Member</span></span> |
| [<span data-ttu-id="9a830-172">to</span><span class="sxs-lookup"><span data-stu-id="9a830-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="9a830-173">Membro</span><span class="sxs-lookup"><span data-stu-id="9a830-173">Member</span></span> |
| [<span data-ttu-id="9a830-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="9a830-175">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-175">Method</span></span> |
| [<span data-ttu-id="9a830-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="9a830-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="9a830-177">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-177">Method</span></span> |
| [<span data-ttu-id="9a830-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="9a830-179">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-179">Method</span></span> |
| [<span data-ttu-id="9a830-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="9a830-181">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-181">Method</span></span> |
| [<span data-ttu-id="9a830-182">close</span><span class="sxs-lookup"><span data-stu-id="9a830-182">close</span></span>](#close) | <span data-ttu-id="9a830-183">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-183">Method</span></span> |
| [<span data-ttu-id="9a830-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="9a830-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="9a830-185">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-185">Method</span></span> |
| [<span data-ttu-id="9a830-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="9a830-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="9a830-187">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-187">Method</span></span> |
| [<span data-ttu-id="9a830-188">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-188">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="9a830-189">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-189">Method</span></span> |
| [<span data-ttu-id="9a830-190">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-190">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="9a830-191">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-191">Method</span></span> |
| [<span data-ttu-id="9a830-192">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-192">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="9a830-193">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-193">Method</span></span> |
| [<span data-ttu-id="9a830-194">getEntities</span><span class="sxs-lookup"><span data-stu-id="9a830-194">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="9a830-195">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-195">Method</span></span> |
| [<span data-ttu-id="9a830-196">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="9a830-196">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="9a830-197">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-197">Method</span></span> |
| [<span data-ttu-id="9a830-198">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="9a830-198">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="9a830-199">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-199">Method</span></span> |
| [<span data-ttu-id="9a830-200">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="9a830-201">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-201">Method</span></span> |
| [<span data-ttu-id="9a830-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="9a830-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="9a830-203">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-203">Method</span></span> |
| [<span data-ttu-id="9a830-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="9a830-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="9a830-205">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-205">Method</span></span> |
| [<span data-ttu-id="9a830-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="9a830-207">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-207">Method</span></span> |
| [<span data-ttu-id="9a830-208">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="9a830-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="9a830-209">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-209">Method</span></span> |
| [<span data-ttu-id="9a830-210">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="9a830-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="9a830-211">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-211">Method</span></span> |
| [<span data-ttu-id="9a830-212">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="9a830-213">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-213">Method</span></span> |
| [<span data-ttu-id="9a830-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="9a830-215">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-215">Method</span></span> |
| [<span data-ttu-id="9a830-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="9a830-217">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-217">Method</span></span> |
| [<span data-ttu-id="9a830-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="9a830-219">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-219">Method</span></span> |
| [<span data-ttu-id="9a830-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="9a830-221">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-221">Method</span></span> |
| [<span data-ttu-id="9a830-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="9a830-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="9a830-223">Método</span><span class="sxs-lookup"><span data-stu-id="9a830-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="9a830-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-224">Example</span></span>

<span data-ttu-id="9a830-225">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="9a830-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="9a830-226">Members</span><span class="sxs-lookup"><span data-stu-id="9a830-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a><span data-ttu-id="9a830-227">anexos: Matriz.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="9a830-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

<span data-ttu-id="9a830-228">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="9a830-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="9a830-229">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9a830-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-230">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="9a830-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="9a830-231">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="9a830-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-232">Type</span></span>

*   <span data-ttu-id="9a830-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="9a830-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-234">Requirements</span></span>

|<span data-ttu-id="9a830-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-235">Requirement</span></span>|<span data-ttu-id="9a830-236">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-238">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-238">1.0</span></span>|
|[<span data-ttu-id="9a830-239">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-240">ReadItem</span></span>|
|[<span data-ttu-id="9a830-241">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-242">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-243">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-243">Example</span></span>

<span data-ttu-id="9a830-244">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="9a830-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="9a830-245">cco :[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-245">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-246">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a830-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="9a830-247">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="9a830-247">Compose mode only.</span></span>

<span data-ttu-id="9a830-248">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-248">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="9a830-249">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="9a830-249">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="9a830-250">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-250">Get 500 members maximum.</span></span>
- <span data-ttu-id="9a830-251">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="9a830-251">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-252">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-252">Type</span></span>

*   [<span data-ttu-id="9a830-253">Destinatários</span><span class="sxs-lookup"><span data-stu-id="9a830-253">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="9a830-254">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-254">Requirements</span></span>

|<span data-ttu-id="9a830-255">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-255">Requirement</span></span>|<span data-ttu-id="9a830-256">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-256">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-257">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-257">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-258">1.1</span><span class="sxs-lookup"><span data-stu-id="9a830-258">1.1</span></span>|
|[<span data-ttu-id="9a830-259">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-259">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-260">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-260">ReadItem</span></span>|
|[<span data-ttu-id="9a830-261">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-261">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-262">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-262">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-263">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-263">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-18"></a><span data-ttu-id="9a830-264">corpo: [Corpo](/javascript/api/outlook/office.body?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-264">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-265">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="9a830-265">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-266">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-266">Type</span></span>

*   [<span data-ttu-id="9a830-267">Body</span><span class="sxs-lookup"><span data-stu-id="9a830-267">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="9a830-268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-268">Requirements</span></span>

|<span data-ttu-id="9a830-269">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-269">Requirement</span></span>|<span data-ttu-id="9a830-270">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-271">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-272">1.1</span><span class="sxs-lookup"><span data-stu-id="9a830-272">1.1</span></span>|
|[<span data-ttu-id="9a830-273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-274">ReadItem</span></span>|
|[<span data-ttu-id="9a830-275">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-276">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-277">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-277">Example</span></span>

<span data-ttu-id="9a830-278">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="9a830-278">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="9a830-279">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-279">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategoriesviewoutlook-js-18"></a><span data-ttu-id="9a830-280">Categorias: [categorias](/javascript/api/outlook/office.categories?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-280">categories: [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-281">Obtém um objeto que fornece métodos para gerenciar as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="9a830-281">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-282">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="9a830-282">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-283">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-283">Type</span></span>

*   [<span data-ttu-id="9a830-284">Categories</span><span class="sxs-lookup"><span data-stu-id="9a830-284">Categories</span></span>](/javascript/api/outlook/office.categories?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="9a830-285">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-285">Requirements</span></span>

|<span data-ttu-id="9a830-286">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-286">Requirement</span></span>|<span data-ttu-id="9a830-287">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-288">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-289">1,8</span><span class="sxs-lookup"><span data-stu-id="9a830-289">1.8</span></span>|
|[<span data-ttu-id="9a830-290">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-291">ReadItem</span></span>|
|[<span data-ttu-id="9a830-292">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-293">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-294">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-294">Example</span></span>

<span data-ttu-id="9a830-295">Este exemplo obtém as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="9a830-295">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="9a830-296">cc : Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-296">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-297">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a830-297">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="9a830-298">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="9a830-298">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a830-299">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9a830-299">Read mode</span></span>

<span data-ttu-id="9a830-300">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a830-300">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="9a830-301">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-301">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="9a830-302">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-302">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="9a830-303">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9a830-303">Compose mode</span></span>

<span data-ttu-id="9a830-304">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a830-304">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="9a830-305">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-305">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="9a830-306">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="9a830-306">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="9a830-307">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-307">Get 500 members maximum.</span></span>
- <span data-ttu-id="9a830-308">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="9a830-308">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9a830-309">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-309">Type</span></span>

*   <span data-ttu-id="9a830-310">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-310">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-311">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-311">Requirements</span></span>

|<span data-ttu-id="9a830-312">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-312">Requirement</span></span>|<span data-ttu-id="9a830-313">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-313">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-314">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-314">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-315">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-315">1.0</span></span>|
|[<span data-ttu-id="9a830-316">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-316">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-317">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-317">ReadItem</span></span>|
|[<span data-ttu-id="9a830-318">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-318">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-319">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-319">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="9a830-320">(anulável) conversationId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-320">(nullable) conversationId: String</span></span>

<span data-ttu-id="9a830-321">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="9a830-321">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="9a830-p109">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="9a830-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="9a830-p110">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="9a830-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-326">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-326">Type</span></span>

*   <span data-ttu-id="9a830-327">String</span><span class="sxs-lookup"><span data-stu-id="9a830-327">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-328">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-328">Requirements</span></span>

|<span data-ttu-id="9a830-329">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-329">Requirement</span></span>|<span data-ttu-id="9a830-330">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-330">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-331">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-331">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-332">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-332">1.0</span></span>|
|[<span data-ttu-id="9a830-333">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-333">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-334">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-334">ReadItem</span></span>|
|[<span data-ttu-id="9a830-335">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-335">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-336">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-336">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-337">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-337">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="9a830-338">dateTimeCreated: Data</span><span class="sxs-lookup"><span data-stu-id="9a830-338">dateTimeCreated: Date</span></span>

<span data-ttu-id="9a830-p111">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="9a830-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-341">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-341">Type</span></span>

*   <span data-ttu-id="9a830-342">Data</span><span class="sxs-lookup"><span data-stu-id="9a830-342">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-343">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-343">Requirements</span></span>

|<span data-ttu-id="9a830-344">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-344">Requirement</span></span>|<span data-ttu-id="9a830-345">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-346">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-347">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-347">1.0</span></span>|
|[<span data-ttu-id="9a830-348">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-349">ReadItem</span></span>|
|[<span data-ttu-id="9a830-350">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-351">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-352">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-352">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="9a830-353">dateTimeModified: Data</span><span class="sxs-lookup"><span data-stu-id="9a830-353">dateTimeModified: Date</span></span>

<span data-ttu-id="9a830-p112">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9a830-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-356">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="9a830-356">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-357">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-357">Type</span></span>

*   <span data-ttu-id="9a830-358">Data</span><span class="sxs-lookup"><span data-stu-id="9a830-358">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-359">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-359">Requirements</span></span>

|<span data-ttu-id="9a830-360">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-360">Requirement</span></span>|<span data-ttu-id="9a830-361">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-362">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-363">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-363">1.0</span></span>|
|[<span data-ttu-id="9a830-364">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-365">ReadItem</span></span>|
|[<span data-ttu-id="9a830-366">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-367">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-367">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-368">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-368">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a><span data-ttu-id="9a830-369">fim: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-369">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-370">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="9a830-370">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="9a830-p113">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="9a830-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a830-373">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9a830-373">Read mode</span></span>

<span data-ttu-id="9a830-374">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="9a830-374">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="9a830-375">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9a830-375">Compose mode</span></span>

<span data-ttu-id="9a830-376">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="9a830-376">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="9a830-377">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="9a830-377">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="9a830-378">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="9a830-378">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="9a830-379">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-379">Type</span></span>

*   <span data-ttu-id="9a830-380">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-380">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-381">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-381">Requirements</span></span>

|<span data-ttu-id="9a830-382">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-382">Requirement</span></span>|<span data-ttu-id="9a830-383">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-383">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-384">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-384">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-385">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-385">1.0</span></span>|
|[<span data-ttu-id="9a830-386">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-386">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-387">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-387">ReadItem</span></span>|
|[<span data-ttu-id="9a830-388">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-388">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-389">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-389">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocationviewoutlook-js-18"></a><span data-ttu-id="9a830-390">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-390">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-391">Obtém ou define os locais de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-391">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a830-392">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9a830-392">Read mode</span></span>

<span data-ttu-id="9a830-393">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) que permite que você obtenha o conjunto de locais (cada um representado por um objeto [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8) ) associado ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-393">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9a830-394">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9a830-394">Compose mode</span></span>

<span data-ttu-id="9a830-395">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) que fornece métodos para obter, remover ou adicionar locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-395">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-396">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-396">Type</span></span>

*   [<span data-ttu-id="9a830-397">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="9a830-397">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="9a830-398">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-398">Requirements</span></span>

|<span data-ttu-id="9a830-399">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-399">Requirement</span></span>|<span data-ttu-id="9a830-400">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-400">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-401">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-401">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-402">1,8</span><span class="sxs-lookup"><span data-stu-id="9a830-402">1.8</span></span>|
|[<span data-ttu-id="9a830-403">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-403">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-404">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-404">ReadItem</span></span>|
|[<span data-ttu-id="9a830-405">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-405">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-406">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-406">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-407">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-407">Example</span></span>

<span data-ttu-id="9a830-408">O exemplo a seguir obtém os locais atuais associados ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-408">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18fromjavascriptapioutlookofficefromviewoutlook-js-18"></a><span data-ttu-id="9a830-409">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[de](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-409">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-410">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a830-410">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="9a830-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="9a830-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-413">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9a830-413">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a830-414">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9a830-414">Read mode</span></span>

<span data-ttu-id="9a830-415">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="9a830-415">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="9a830-416">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9a830-416">Compose mode</span></span>

<span data-ttu-id="9a830-417">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="9a830-417">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9a830-418">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-418">Type</span></span>

*   <span data-ttu-id="9a830-419">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [de](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-419">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-420">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-420">Requirements</span></span>

|<span data-ttu-id="9a830-421">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-421">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="9a830-422">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-423">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-423">1.0</span></span>|<span data-ttu-id="9a830-424">1.7</span><span class="sxs-lookup"><span data-stu-id="9a830-424">1.7</span></span>|
|[<span data-ttu-id="9a830-425">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-426">ReadItem</span></span>|<span data-ttu-id="9a830-427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a830-427">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a830-428">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-429">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-429">Read</span></span>|<span data-ttu-id="9a830-430">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-430">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheadersviewoutlook-js-18"></a><span data-ttu-id="9a830-431">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-431">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-432">Obtém ou define cabeçalhos de Internet personalizados em uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a830-432">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="9a830-433">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="9a830-433">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-434">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-434">Type</span></span>

*   [<span data-ttu-id="9a830-435">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="9a830-435">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="9a830-436">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-436">Requirements</span></span>

|<span data-ttu-id="9a830-437">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-437">Requirement</span></span>|<span data-ttu-id="9a830-438">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-438">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-439">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-440">1,8</span><span class="sxs-lookup"><span data-stu-id="9a830-440">1.8</span></span>|
|[<span data-ttu-id="9a830-441">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-441">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-442">ReadItem</span></span>|
|[<span data-ttu-id="9a830-443">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-443">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-444">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-444">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-445">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-445">Example</span></span>

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

#### <a name="internetmessageid-string"></a><span data-ttu-id="9a830-446">internetMessageId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-446">internetMessageId: String</span></span>

<span data-ttu-id="9a830-p116">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="9a830-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-449">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-449">Type</span></span>

*   <span data-ttu-id="9a830-450">String</span><span class="sxs-lookup"><span data-stu-id="9a830-450">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-451">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-451">Requirements</span></span>

|<span data-ttu-id="9a830-452">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-452">Requirement</span></span>|<span data-ttu-id="9a830-453">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-453">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-454">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-454">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-455">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-455">1.0</span></span>|
|[<span data-ttu-id="9a830-456">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-456">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-457">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-457">ReadItem</span></span>|
|[<span data-ttu-id="9a830-458">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-458">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-459">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-459">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-460">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-460">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="9a830-461">itemClass: Cadeia de caracteres </span><span class="sxs-lookup"><span data-stu-id="9a830-461">itemClass: String</span></span>

<span data-ttu-id="9a830-p117">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9a830-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="9a830-p118">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="9a830-466">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-466">Type</span></span>|<span data-ttu-id="9a830-467">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-467">Description</span></span>|<span data-ttu-id="9a830-468">classe de item</span><span class="sxs-lookup"><span data-stu-id="9a830-468">item class</span></span>|
|---|---|---|
|<span data-ttu-id="9a830-469">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="9a830-469">Appointment items</span></span>|<span data-ttu-id="9a830-470">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="9a830-470">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="9a830-471">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="9a830-471">Message items</span></span>|<span data-ttu-id="9a830-472">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="9a830-472">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="9a830-473">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="9a830-473">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-474">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-474">Type</span></span>

*   <span data-ttu-id="9a830-475">String</span><span class="sxs-lookup"><span data-stu-id="9a830-475">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-476">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-476">Requirements</span></span>

|<span data-ttu-id="9a830-477">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-477">Requirement</span></span>|<span data-ttu-id="9a830-478">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-478">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-479">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-479">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-480">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-480">1.0</span></span>|
|[<span data-ttu-id="9a830-481">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-481">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-482">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-482">ReadItem</span></span>|
|[<span data-ttu-id="9a830-483">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-483">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-484">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-484">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-485">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-485">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="9a830-486">(anulável) itemId: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-486">(nullable) itemId: String</span></span>

<span data-ttu-id="9a830-p119">Obtém o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9a830-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-489">O identificador retornado pela propriedade `itemId` é o mesmo que o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="9a830-489">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="9a830-490">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="9a830-490">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="9a830-491">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="9a830-491">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="9a830-492">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="9a830-492">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="9a830-p121">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-495">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-495">Type</span></span>

*   <span data-ttu-id="9a830-496">String</span><span class="sxs-lookup"><span data-stu-id="9a830-496">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-497">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-497">Requirements</span></span>

|<span data-ttu-id="9a830-498">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-498">Requirement</span></span>|<span data-ttu-id="9a830-499">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-499">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-500">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-500">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-501">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-501">1.0</span></span>|
|[<span data-ttu-id="9a830-502">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-502">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-503">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-503">ReadItem</span></span>|
|[<span data-ttu-id="9a830-504">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-504">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-505">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-505">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-506">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-506">Example</span></span>

<span data-ttu-id="9a830-p122">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="9a830-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-18"></a><span data-ttu-id="9a830-509">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-509">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-510">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="9a830-510">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="9a830-511">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-511">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-512">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-512">Type</span></span>

*   [<span data-ttu-id="9a830-513">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="9a830-513">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="9a830-514">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-514">Requirements</span></span>

|<span data-ttu-id="9a830-515">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-515">Requirement</span></span>|<span data-ttu-id="9a830-516">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-516">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-517">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-517">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-518">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-518">1.0</span></span>|
|[<span data-ttu-id="9a830-519">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-519">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-520">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-520">ReadItem</span></span>|
|[<span data-ttu-id="9a830-521">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-521">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-522">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-522">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-523">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-523">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-18"></a><span data-ttu-id="9a830-524">Local: Cadeia de caracteres[Local](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-524">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-525">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-525">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a830-526">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9a830-526">Read mode</span></span>

<span data-ttu-id="9a830-527">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-527">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="9a830-528">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9a830-528">Compose mode</span></span>

<span data-ttu-id="9a830-529">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-529">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9a830-530">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-530">Type</span></span>

*   <span data-ttu-id="9a830-531">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-531">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-532">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-532">Requirements</span></span>

|<span data-ttu-id="9a830-533">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-533">Requirement</span></span>|<span data-ttu-id="9a830-534">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-534">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-535">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-535">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-536">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-536">1.0</span></span>|
|[<span data-ttu-id="9a830-537">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-537">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-538">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-538">ReadItem</span></span>|
|[<span data-ttu-id="9a830-539">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-539">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-540">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-540">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="9a830-541">normalizedSubject: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-541">normalizedSubject: String</span></span>

<span data-ttu-id="9a830-p123">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9a830-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="9a830-p124">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="9a830-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-546">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-546">Type</span></span>

*   <span data-ttu-id="9a830-547">String</span><span class="sxs-lookup"><span data-stu-id="9a830-547">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-548">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-548">Requirements</span></span>

|<span data-ttu-id="9a830-549">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-549">Requirement</span></span>|<span data-ttu-id="9a830-550">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-551">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-552">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-552">1.0</span></span>|
|[<span data-ttu-id="9a830-553">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-553">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-554">ReadItem</span></span>|
|[<span data-ttu-id="9a830-555">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-555">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-556">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-556">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-557">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-557">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-18"></a><span data-ttu-id="9a830-558">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-558">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-559">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="9a830-559">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-560">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-560">Type</span></span>

*   [<span data-ttu-id="9a830-561">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="9a830-561">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="9a830-562">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-562">Requirements</span></span>

|<span data-ttu-id="9a830-563">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-563">Requirement</span></span>|<span data-ttu-id="9a830-564">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-565">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-566">1.3</span><span class="sxs-lookup"><span data-stu-id="9a830-566">1.3</span></span>|
|[<span data-ttu-id="9a830-567">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-567">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-568">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-568">ReadItem</span></span>|
|[<span data-ttu-id="9a830-569">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-569">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-570">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-570">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-571">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-571">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="9a830-572">optionalAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-572">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-573">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="9a830-573">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="9a830-574">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="9a830-574">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a830-575">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9a830-575">Read mode</span></span>

<span data-ttu-id="9a830-576">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="9a830-576">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="9a830-577">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-577">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="9a830-578">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-578">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="9a830-579">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9a830-579">Compose mode</span></span>

<span data-ttu-id="9a830-580">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="9a830-580">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="9a830-581">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-581">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="9a830-582">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="9a830-582">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="9a830-583">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-583">Get 500 members maximum.</span></span>
- <span data-ttu-id="9a830-584">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="9a830-584">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9a830-585">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-585">Type</span></span>

*   <span data-ttu-id="9a830-586">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-586">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-587">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-587">Requirements</span></span>

|<span data-ttu-id="9a830-588">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-588">Requirement</span></span>|<span data-ttu-id="9a830-589">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-589">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-590">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-591">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-591">1.0</span></span>|
|[<span data-ttu-id="9a830-592">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-592">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-593">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-593">ReadItem</span></span>|
|[<span data-ttu-id="9a830-594">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-594">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-595">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-595">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18organizerjavascriptapioutlookofficeorganizerviewoutlook-js-18"></a><span data-ttu-id="9a830-596">organizador: [](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9a830-596">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-597">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="9a830-597">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a830-598">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9a830-598">Read mode</span></span>

<span data-ttu-id="9a830-599">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="9a830-599">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="9a830-600">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9a830-600">Compose mode</span></span>

<span data-ttu-id="9a830-601">A `organizer` propriedade retorna um objeto [organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="9a830-601">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="9a830-602">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-602">Type</span></span>

*   <span data-ttu-id="9a830-603">[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [Organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9a830-603">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-604">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-604">Requirements</span></span>

|<span data-ttu-id="9a830-605">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-605">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="9a830-606">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-607">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-607">1.0</span></span>|<span data-ttu-id="9a830-608">1.7</span><span class="sxs-lookup"><span data-stu-id="9a830-608">1.7</span></span>|
|[<span data-ttu-id="9a830-609">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-610">ReadItem</span></span>|<span data-ttu-id="9a830-611">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a830-611">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a830-612">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-613">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-613">Read</span></span>|<span data-ttu-id="9a830-614">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-614">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-18"></a><span data-ttu-id="9a830-615">(anulável) recorrência: [recorrência](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-615">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-616">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-616">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="9a830-617">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="9a830-617">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="9a830-618">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-618">Read and compose modes for appointment items.</span></span> <span data-ttu-id="9a830-619">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="9a830-619">Read mode for meeting request items.</span></span>

<span data-ttu-id="9a830-620">A `recurrence` propriedade retorna um objeto de [recorrência](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="9a830-620">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="9a830-621">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="9a830-621">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="9a830-622">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="9a830-622">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="9a830-623">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="9a830-623">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="9a830-624">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="9a830-624">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a830-625">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9a830-625">Read mode</span></span>

<span data-ttu-id="9a830-626">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-626">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object that represents the appointment recurrence.</span></span> <span data-ttu-id="9a830-627">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="9a830-627">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="9a830-628">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9a830-628">Compose mode</span></span>

<span data-ttu-id="9a830-629">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-629">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="9a830-630">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="9a830-630">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="9a830-631">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-631">Type</span></span>

* [<span data-ttu-id="9a830-632">Recorrência</span><span class="sxs-lookup"><span data-stu-id="9a830-632">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)

|<span data-ttu-id="9a830-633">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-633">Requirement</span></span>|<span data-ttu-id="9a830-634">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-635">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-636">1.7</span><span class="sxs-lookup"><span data-stu-id="9a830-636">1.7</span></span>|
|[<span data-ttu-id="9a830-637">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-638">ReadItem</span></span>|
|[<span data-ttu-id="9a830-639">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-640">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-640">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="9a830-641">requiredAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-641">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-642">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="9a830-642">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="9a830-643">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="9a830-643">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a830-644">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9a830-644">Read mode</span></span>

<span data-ttu-id="9a830-645">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="9a830-645">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="9a830-646">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-646">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="9a830-647">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-647">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="9a830-648">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9a830-648">Compose mode</span></span>

<span data-ttu-id="9a830-649">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="9a830-649">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="9a830-650">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-650">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="9a830-651">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="9a830-651">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="9a830-652">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-652">Get 500 members maximum.</span></span>
- <span data-ttu-id="9a830-653">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="9a830-653">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="9a830-654">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-654">Type</span></span>

*   <span data-ttu-id="9a830-655">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-655">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-656">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-656">Requirements</span></span>

|<span data-ttu-id="9a830-657">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-657">Requirement</span></span>|<span data-ttu-id="9a830-658">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-658">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-659">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-659">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-660">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-660">1.0</span></span>|
|[<span data-ttu-id="9a830-661">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-661">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-662">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-662">ReadItem</span></span>|
|[<span data-ttu-id="9a830-663">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-663">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-664">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-664">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18"></a><span data-ttu-id="9a830-665">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-665">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-p135">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9a830-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="9a830-p136">As propriedades [`from`](#from-emailaddressdetailsfrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="9a830-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-670">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9a830-670">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-671">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-671">Type</span></span>

*   [<span data-ttu-id="9a830-672">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9a830-672">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="9a830-673">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-673">Requirements</span></span>

|<span data-ttu-id="9a830-674">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-674">Requirement</span></span>|<span data-ttu-id="9a830-675">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-675">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-676">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-676">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-677">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-677">1.0</span></span>|
|[<span data-ttu-id="9a830-678">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-678">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-679">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-679">ReadItem</span></span>|
|[<span data-ttu-id="9a830-680">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-680">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-681">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-681">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-682">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-682">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="9a830-683">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="9a830-683">(nullable) seriesId: String</span></span>

<span data-ttu-id="9a830-684">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="9a830-684">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="9a830-685">No Outlook na Web e clientes de desktop, o `seriesId` retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="9a830-685">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="9a830-686">No entanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="9a830-686">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-687">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="9a830-687">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="9a830-688">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="9a830-688">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="9a830-689">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="9a830-689">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="9a830-690">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="9a830-690">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="9a830-691">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="9a830-691">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="9a830-692">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-692">Type</span></span>

* <span data-ttu-id="9a830-693">String</span><span class="sxs-lookup"><span data-stu-id="9a830-693">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-694">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-694">Requirements</span></span>

|<span data-ttu-id="9a830-695">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-695">Requirement</span></span>|<span data-ttu-id="9a830-696">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-697">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-698">1.7</span><span class="sxs-lookup"><span data-stu-id="9a830-698">1.7</span></span>|
|[<span data-ttu-id="9a830-699">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-700">ReadItem</span></span>|
|[<span data-ttu-id="9a830-701">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-702">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-702">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-703">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-703">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a><span data-ttu-id="9a830-704">início: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-704">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-705">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="9a830-705">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="9a830-p139">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="9a830-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a830-708">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9a830-708">Read mode</span></span>

<span data-ttu-id="9a830-709">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="9a830-709">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="9a830-710">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9a830-710">Compose mode</span></span>

<span data-ttu-id="9a830-711">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="9a830-711">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="9a830-712">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="9a830-712">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="9a830-713">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="9a830-713">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="9a830-714">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-714">Type</span></span>

*   <span data-ttu-id="9a830-715">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-715">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-716">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-716">Requirements</span></span>

|<span data-ttu-id="9a830-717">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-717">Requirement</span></span>|<span data-ttu-id="9a830-718">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-719">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-719">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-720">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-720">1.0</span></span>|
|[<span data-ttu-id="9a830-721">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-721">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-722">ReadItem</span></span>|
|[<span data-ttu-id="9a830-723">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-723">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-724">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-724">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-18"></a><span data-ttu-id="9a830-725">Assunto: Cadeia de caracteres|[Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-725">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-726">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="9a830-726">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="9a830-727">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="9a830-727">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a830-728">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9a830-728">Read mode</span></span>

<span data-ttu-id="9a830-p140">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="9a830-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="9a830-731">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="9a830-731">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="9a830-732">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9a830-732">Compose mode</span></span>
<span data-ttu-id="9a830-733">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="9a830-733">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="9a830-734">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-734">Type</span></span>

*   <span data-ttu-id="9a830-735">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-735">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-736">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-736">Requirements</span></span>

|<span data-ttu-id="9a830-737">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-737">Requirement</span></span>|<span data-ttu-id="9a830-738">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-738">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-739">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-739">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-740">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-740">1.0</span></span>|
|[<span data-ttu-id="9a830-741">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-741">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-742">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-742">ReadItem</span></span>|
|[<span data-ttu-id="9a830-743">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-743">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-744">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-744">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="9a830-745">para: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-745">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-746">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a830-746">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="9a830-747">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="9a830-747">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a830-748">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="9a830-748">Read mode</span></span>

<span data-ttu-id="9a830-749">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a830-749">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="9a830-750">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-750">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="9a830-751">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-751">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="9a830-752">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="9a830-752">Compose mode</span></span>

<span data-ttu-id="9a830-753">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a830-753">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="9a830-754">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-754">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="9a830-755">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="9a830-755">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="9a830-756">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="9a830-756">Get 500 members maximum.</span></span>
- <span data-ttu-id="9a830-757">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="9a830-757">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9a830-758">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-758">Type</span></span>

*   <span data-ttu-id="9a830-759">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-759">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-760">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-760">Requirements</span></span>

|<span data-ttu-id="9a830-761">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-761">Requirement</span></span>|<span data-ttu-id="9a830-762">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-762">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-763">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-763">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-764">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-764">1.0</span></span>|
|[<span data-ttu-id="9a830-765">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-765">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-766">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-766">ReadItem</span></span>|
|[<span data-ttu-id="9a830-767">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-767">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-768">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-768">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="9a830-769">Métodos</span><span class="sxs-lookup"><span data-stu-id="9a830-769">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="9a830-770">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a830-770">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9a830-771">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="9a830-771">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="9a830-772">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="9a830-772">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="9a830-773">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="9a830-773">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-774">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-774">Parameters</span></span>
|<span data-ttu-id="9a830-775">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-775">Name</span></span>|<span data-ttu-id="9a830-776">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-776">Type</span></span>|<span data-ttu-id="9a830-777">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-777">Attributes</span></span>|<span data-ttu-id="9a830-778">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-778">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="9a830-779">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-779">String</span></span>||<span data-ttu-id="9a830-p144">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="9a830-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="9a830-782">String</span><span class="sxs-lookup"><span data-stu-id="9a830-782">String</span></span>||<span data-ttu-id="9a830-p145">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="9a830-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="9a830-785">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-785">Object</span></span>|<span data-ttu-id="9a830-786">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-786">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-787">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-787">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a830-788">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-788">Object</span></span>|<span data-ttu-id="9a830-789">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-789">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-790">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-790">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="9a830-791">Booliano</span><span class="sxs-lookup"><span data-stu-id="9a830-791">Boolean</span></span>|<span data-ttu-id="9a830-792">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-792">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-793">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="9a830-793">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="9a830-794">function</span><span class="sxs-lookup"><span data-stu-id="9a830-794">function</span></span>|<span data-ttu-id="9a830-795">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-795">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-796">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-796">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9a830-797">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9a830-797">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9a830-798">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="9a830-798">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9a830-799">Erros</span><span class="sxs-lookup"><span data-stu-id="9a830-799">Errors</span></span>

|<span data-ttu-id="9a830-800">Código de erro</span><span class="sxs-lookup"><span data-stu-id="9a830-800">Error code</span></span>|<span data-ttu-id="9a830-801">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-801">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="9a830-802">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="9a830-802">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="9a830-803">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="9a830-803">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="9a830-804">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="9a830-804">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-805">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-805">Requirements</span></span>

|<span data-ttu-id="9a830-806">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-806">Requirement</span></span>|<span data-ttu-id="9a830-807">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-808">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-809">1.1</span><span class="sxs-lookup"><span data-stu-id="9a830-809">1.1</span></span>|
|[<span data-ttu-id="9a830-810">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-811">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a830-811">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a830-812">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-813">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-813">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a830-814">Exemplos</span><span class="sxs-lookup"><span data-stu-id="9a830-814">Examples</span></span>

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

<span data-ttu-id="9a830-815">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a830-815">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="9a830-816">addFileAttachmentFromBase64Async (base64file, AttachmentName, [Options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a830-816">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9a830-817">Adiciona um arquivo da codificação Base64 a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="9a830-817">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="9a830-818">O `addFileAttachmentFromBase64Async` método carrega o arquivo da codificação Base64 e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="9a830-818">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="9a830-819">Esse método retorna o identificador de anexo no objeto AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="9a830-819">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="9a830-820">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="9a830-820">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-821">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-821">Parameters</span></span>

|<span data-ttu-id="9a830-822">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-822">Name</span></span>|<span data-ttu-id="9a830-823">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-823">Type</span></span>|<span data-ttu-id="9a830-824">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-824">Attributes</span></span>|<span data-ttu-id="9a830-825">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-825">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="9a830-826">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-826">String</span></span>||<span data-ttu-id="9a830-827">O conteúdo codificado em Base64 de uma imagem ou arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="9a830-827">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="9a830-828">String</span><span class="sxs-lookup"><span data-stu-id="9a830-828">String</span></span>||<span data-ttu-id="9a830-p147">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="9a830-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="9a830-831">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-831">Object</span></span>|<span data-ttu-id="9a830-832">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-832">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-833">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-833">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a830-834">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-834">Object</span></span>|<span data-ttu-id="9a830-835">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-835">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-836">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-836">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="9a830-837">Booliano</span><span class="sxs-lookup"><span data-stu-id="9a830-837">Boolean</span></span>|<span data-ttu-id="9a830-838">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-838">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-839">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="9a830-839">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="9a830-840">function</span><span class="sxs-lookup"><span data-stu-id="9a830-840">function</span></span>|<span data-ttu-id="9a830-841">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-841">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-842">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-842">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9a830-843">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9a830-843">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9a830-844">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="9a830-844">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9a830-845">Erros</span><span class="sxs-lookup"><span data-stu-id="9a830-845">Errors</span></span>

|<span data-ttu-id="9a830-846">Código de erro</span><span class="sxs-lookup"><span data-stu-id="9a830-846">Error code</span></span>|<span data-ttu-id="9a830-847">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-847">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="9a830-848">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="9a830-848">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="9a830-849">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="9a830-849">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="9a830-850">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="9a830-850">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-851">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-851">Requirements</span></span>

|<span data-ttu-id="9a830-852">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-852">Requirement</span></span>|<span data-ttu-id="9a830-853">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-853">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-854">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-854">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-855">1,8</span><span class="sxs-lookup"><span data-stu-id="9a830-855">1.8</span></span>|
|[<span data-ttu-id="9a830-856">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-856">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-857">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a830-857">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a830-858">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-858">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-859">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-859">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a830-860">Exemplos</span><span class="sxs-lookup"><span data-stu-id="9a830-860">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="9a830-861">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a830-861">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="9a830-862">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="9a830-862">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="9a830-863">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="9a830-863">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-864">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-864">Parameters</span></span>

| <span data-ttu-id="9a830-865">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-865">Name</span></span> | <span data-ttu-id="9a830-866">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-866">Type</span></span> | <span data-ttu-id="9a830-867">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-867">Attributes</span></span> | <span data-ttu-id="9a830-868">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-868">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="9a830-869">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="9a830-869">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="9a830-870">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="9a830-870">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="9a830-871">Função</span><span class="sxs-lookup"><span data-stu-id="9a830-871">Function</span></span> || <span data-ttu-id="9a830-p148">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="9a830-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="9a830-875">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-875">Object</span></span> | <span data-ttu-id="9a830-876">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-876">&lt;optional&gt;</span></span> | <span data-ttu-id="9a830-877">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-877">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9a830-878">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-878">Object</span></span> | <span data-ttu-id="9a830-879">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-879">&lt;optional&gt;</span></span> | <span data-ttu-id="9a830-880">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-880">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="9a830-881">function</span><span class="sxs-lookup"><span data-stu-id="9a830-881">function</span></span>| <span data-ttu-id="9a830-882">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-882">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-883">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-884">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-884">Requirements</span></span>

|<span data-ttu-id="9a830-885">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-885">Requirement</span></span>| <span data-ttu-id="9a830-886">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-887">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a830-888">1.7</span><span class="sxs-lookup"><span data-stu-id="9a830-888">1.7</span></span> |
|[<span data-ttu-id="9a830-889">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a830-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-890">ReadItem</span></span> |
|[<span data-ttu-id="9a830-891">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9a830-892">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-892">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="9a830-893">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-893">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="9a830-894">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a830-894">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9a830-895">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-895">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="9a830-p149">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="9a830-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="9a830-899">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="9a830-899">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="9a830-900">Se o Suplemento do Office estiver em execução no Outlook na Web, o método `addItemAttachmentAsync` pode anexar itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="9a830-900">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-901">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-901">Parameters</span></span>

|<span data-ttu-id="9a830-902">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-902">Name</span></span>|<span data-ttu-id="9a830-903">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-903">Type</span></span>|<span data-ttu-id="9a830-904">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-904">Attributes</span></span>|<span data-ttu-id="9a830-905">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-905">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="9a830-906">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-906">String</span></span>||<span data-ttu-id="9a830-p150">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="9a830-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="9a830-909">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-909">String</span></span>||<span data-ttu-id="9a830-910">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="9a830-910">The subject of the item to be attached.</span></span> <span data-ttu-id="9a830-911">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="9a830-911">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="9a830-912">Object</span><span class="sxs-lookup"><span data-stu-id="9a830-912">Object</span></span>|<span data-ttu-id="9a830-913">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-913">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-914">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-914">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a830-915">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-915">Object</span></span>|<span data-ttu-id="9a830-916">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-916">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-917">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-917">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a830-918">function</span><span class="sxs-lookup"><span data-stu-id="9a830-918">function</span></span>|<span data-ttu-id="9a830-919">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-919">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-920">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-920">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9a830-921">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9a830-921">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9a830-922">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="9a830-922">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9a830-923">Erros</span><span class="sxs-lookup"><span data-stu-id="9a830-923">Errors</span></span>

|<span data-ttu-id="9a830-924">Código de erro</span><span class="sxs-lookup"><span data-stu-id="9a830-924">Error code</span></span>|<span data-ttu-id="9a830-925">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-925">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="9a830-926">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="9a830-926">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-927">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-927">Requirements</span></span>

|<span data-ttu-id="9a830-928">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-928">Requirement</span></span>|<span data-ttu-id="9a830-929">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-929">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-930">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-930">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-931">1.1</span><span class="sxs-lookup"><span data-stu-id="9a830-931">1.1</span></span>|
|[<span data-ttu-id="9a830-932">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-932">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-933">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a830-933">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a830-934">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-934">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-935">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-935">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-936">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-936">Example</span></span>

<span data-ttu-id="9a830-937">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="9a830-937">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="9a830-938">close()</span><span class="sxs-lookup"><span data-stu-id="9a830-938">close()</span></span>

<span data-ttu-id="9a830-939">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="9a830-939">Closes the current item that is being composed.</span></span>

<span data-ttu-id="9a830-p152">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="9a830-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-942">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="9a830-942">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="9a830-943">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="9a830-943">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-944">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-944">Requirements</span></span>

|<span data-ttu-id="9a830-945">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-945">Requirement</span></span>|<span data-ttu-id="9a830-946">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-947">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-948">1.3</span><span class="sxs-lookup"><span data-stu-id="9a830-948">1.3</span></span>|
|[<span data-ttu-id="9a830-949">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-950">Restrito</span><span class="sxs-lookup"><span data-stu-id="9a830-950">Restricted</span></span>|
|[<span data-ttu-id="9a830-951">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-952">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-952">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="9a830-953">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="9a830-953">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="9a830-954">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="9a830-954">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-955">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="9a830-955">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9a830-956">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="9a830-956">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9a830-957">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="9a830-957">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="9a830-p153">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="9a830-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-961">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-961">Parameters</span></span>

|<span data-ttu-id="9a830-962">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-962">Name</span></span>|<span data-ttu-id="9a830-963">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-963">Type</span></span>|<span data-ttu-id="9a830-964">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-964">Attributes</span></span>|<span data-ttu-id="9a830-965">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-965">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="9a830-966">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="9a830-966">String &#124; Object</span></span>||<span data-ttu-id="9a830-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="9a830-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9a830-969">**OU**</span><span class="sxs-lookup"><span data-stu-id="9a830-969">**OR**</span></span><br/><span data-ttu-id="9a830-p155">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="9a830-972">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-972">String</span></span>|<span data-ttu-id="9a830-973">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-973">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-p156">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="9a830-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="9a830-976">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-976">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="9a830-977">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-977">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-978">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="9a830-978">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="9a830-979">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-979">String</span></span>||<span data-ttu-id="9a830-p157">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="9a830-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="9a830-982">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-982">String</span></span>||<span data-ttu-id="9a830-983">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="9a830-983">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="9a830-984">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-984">String</span></span>||<span data-ttu-id="9a830-p158">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="9a830-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="9a830-987">Booliano</span><span class="sxs-lookup"><span data-stu-id="9a830-987">Boolean</span></span>||<span data-ttu-id="9a830-p159">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="9a830-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="9a830-990">String</span><span class="sxs-lookup"><span data-stu-id="9a830-990">String</span></span>||<span data-ttu-id="9a830-p160">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="9a830-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="9a830-994">function</span><span class="sxs-lookup"><span data-stu-id="9a830-994">function</span></span>|<span data-ttu-id="9a830-995">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-995">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-996">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-996">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-997">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-997">Requirements</span></span>

|<span data-ttu-id="9a830-998">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-998">Requirement</span></span>|<span data-ttu-id="9a830-999">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-999">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1000">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1000">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1001">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-1001">1.0</span></span>|
|[<span data-ttu-id="9a830-1002">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1002">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1003">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1003">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1004">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-1004">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1005">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-1005">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a830-1006">Exemplos</span><span class="sxs-lookup"><span data-stu-id="9a830-1006">Examples</span></span>

<span data-ttu-id="9a830-1007">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="9a830-1007">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="9a830-1008">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="9a830-1008">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="9a830-1009">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="9a830-1009">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9a830-1010">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="9a830-1010">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="9a830-1011">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="9a830-1011">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="9a830-1012">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1012">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="9a830-1013">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="9a830-1013">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="9a830-1014">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="9a830-1014">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-1015">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="9a830-1015">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9a830-1016">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="9a830-1016">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9a830-1017">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="9a830-1017">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="9a830-p161">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="9a830-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1021">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1021">Parameters</span></span>

|<span data-ttu-id="9a830-1022">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1022">Name</span></span>|<span data-ttu-id="9a830-1023">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1023">Type</span></span>|<span data-ttu-id="9a830-1024">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-1024">Attributes</span></span>|<span data-ttu-id="9a830-1025">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1025">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="9a830-1026">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="9a830-1026">String &#124; Object</span></span>||<span data-ttu-id="9a830-p162">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="9a830-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9a830-1029">**OU**</span><span class="sxs-lookup"><span data-stu-id="9a830-1029">**OR**</span></span><br/><span data-ttu-id="9a830-p163">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="9a830-1032">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-1032">String</span></span>|<span data-ttu-id="9a830-1033">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1033">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-p164">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="9a830-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="9a830-1036">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1036">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="9a830-1037">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1038">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="9a830-1038">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="9a830-1039">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-1039">String</span></span>||<span data-ttu-id="9a830-p165">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="9a830-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="9a830-1042">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-1042">String</span></span>||<span data-ttu-id="9a830-1043">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="9a830-1043">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="9a830-1044">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-1044">String</span></span>||<span data-ttu-id="9a830-p166">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="9a830-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="9a830-1047">Booliano</span><span class="sxs-lookup"><span data-stu-id="9a830-1047">Boolean</span></span>||<span data-ttu-id="9a830-p167">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="9a830-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="9a830-1050">String</span><span class="sxs-lookup"><span data-stu-id="9a830-1050">String</span></span>||<span data-ttu-id="9a830-p168">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="9a830-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="9a830-1054">function</span><span class="sxs-lookup"><span data-stu-id="9a830-1054">function</span></span>|<span data-ttu-id="9a830-1055">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1055">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1056">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-1056">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1057">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1057">Requirements</span></span>

|<span data-ttu-id="9a830-1058">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1058">Requirement</span></span>|<span data-ttu-id="9a830-1059">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1059">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1060">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1060">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1061">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-1061">1.0</span></span>|
|[<span data-ttu-id="9a830-1062">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1062">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1063">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1063">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1064">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1064">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1065">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-1065">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a830-1066">Exemplos</span><span class="sxs-lookup"><span data-stu-id="9a830-1066">Examples</span></span>

<span data-ttu-id="9a830-1067">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="9a830-1067">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="9a830-1068">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="9a830-1068">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="9a830-1069">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="9a830-1069">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9a830-1070">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="9a830-1070">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="9a830-1071">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="9a830-1071">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="9a830-1072">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1072">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="9a830-1073">getAllInternetHeadersAsync ([opções], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a830-1073">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="9a830-1074">Obtém todos os cabeçalhos de Internet da mensagem como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="9a830-1074">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="9a830-1075">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="9a830-1075">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1076">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1076">Parameters</span></span>

|<span data-ttu-id="9a830-1077">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1077">Name</span></span>|<span data-ttu-id="9a830-1078">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1078">Type</span></span>|<span data-ttu-id="9a830-1079">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-1079">Attributes</span></span>|<span data-ttu-id="9a830-1080">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1080">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="9a830-1081">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1081">Object</span></span>|<span data-ttu-id="9a830-1082">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1082">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1083">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-1083">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a830-1084">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1084">Object</span></span>|<span data-ttu-id="9a830-1085">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1085">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1086">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1086">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a830-1087">function</span><span class="sxs-lookup"><span data-stu-id="9a830-1087">function</span></span>|<span data-ttu-id="9a830-1088">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1088">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1089">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="9a830-1090">Com êxito, os dados de cabeçalhos de Internet são fornecidos na propriedade asyncResult. Value como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="9a830-1090">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="9a830-1091">Consulte [RFC 2183](https://tools.ietf.org/html/rfc2183) para obter as informações de formatação do valor de cadeia de caracteres retornado.</span><span class="sxs-lookup"><span data-stu-id="9a830-1091">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="9a830-1092">Se a chamada falhar, a propriedade asyncResult. Error conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="9a830-1092">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1093">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1093">Requirements</span></span>

|<span data-ttu-id="9a830-1094">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1094">Requirement</span></span>|<span data-ttu-id="9a830-1095">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1096">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1097">1,8</span><span class="sxs-lookup"><span data-stu-id="9a830-1097">1.8</span></span>|
|[<span data-ttu-id="9a830-1098">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1098">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1099">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1100">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1100">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1101">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-1101">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a830-1102">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9a830-1102">Returns:</span></span>

<span data-ttu-id="9a830-1103">A Internet cabeçalhos dados como uma cadeia de caracteres formatada de acordo com a [RFC 2183](https://tools.ietf.org/html/rfc2183).</span><span class="sxs-lookup"><span data-stu-id="9a830-1103">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="9a830-1104">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="9a830-1104">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="9a830-1105">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1105">Example</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontentviewoutlook-js-18"></a><span data-ttu-id="9a830-1106">getAttachmentContentAsync (attachmentid, [opções], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-1106">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span></span>

<span data-ttu-id="9a830-1107">Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um `AttachmentContent` objeto.</span><span class="sxs-lookup"><span data-stu-id="9a830-1107">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="9a830-1108">O `getAttachmentContentAsync` método obtém o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="9a830-1108">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="9a830-1109">Como prática recomendada, você deve usar o identificador para recuperar um anexo na mesma sessão em que o attachmentIds foi recuperado com a `getAttachmentsAsync` chamada ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="9a830-1109">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="9a830-1110">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="9a830-1110">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="9a830-1111">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1111">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1112">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1112">Parameters</span></span>

|<span data-ttu-id="9a830-1113">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1113">Name</span></span>|<span data-ttu-id="9a830-1114">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1114">Type</span></span>|<span data-ttu-id="9a830-1115">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-1115">Attributes</span></span>|<span data-ttu-id="9a830-1116">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1116">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="9a830-1117">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-1117">String</span></span>||<span data-ttu-id="9a830-1118">O identificador do anexo que você deseja obter.</span><span class="sxs-lookup"><span data-stu-id="9a830-1118">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="9a830-1119">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1119">Object</span></span>|<span data-ttu-id="9a830-1120">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1120">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1121">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-1121">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a830-1122">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1122">Object</span></span>|<span data-ttu-id="9a830-1123">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1123">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1124">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1124">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a830-1125">function</span><span class="sxs-lookup"><span data-stu-id="9a830-1125">function</span></span>|<span data-ttu-id="9a830-1126">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1126">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1127">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-1127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1128">Requirements</span></span>

|<span data-ttu-id="9a830-1129">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1129">Requirement</span></span>|<span data-ttu-id="9a830-1130">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1130">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1132">1,8</span><span class="sxs-lookup"><span data-stu-id="9a830-1132">1.8</span></span>|
|[<span data-ttu-id="9a830-1133">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1133">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1134">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1135">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-1135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1136">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-1136">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a830-1137">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9a830-1137">Returns:</span></span>

<span data-ttu-id="9a830-1138">Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-1138">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="9a830-1139">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1139">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a><span data-ttu-id="9a830-1140">getAttachmentsAsync ([Options], [callback]) → array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="9a830-1140">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

<span data-ttu-id="9a830-1141">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="9a830-1141">Gets the item's attachments as an array.</span></span> <span data-ttu-id="9a830-1142">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="9a830-1142">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1143">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1143">Parameters</span></span>

|<span data-ttu-id="9a830-1144">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1144">Name</span></span>|<span data-ttu-id="9a830-1145">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1145">Type</span></span>|<span data-ttu-id="9a830-1146">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-1146">Attributes</span></span>|<span data-ttu-id="9a830-1147">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1147">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="9a830-1148">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1148">Object</span></span>|<span data-ttu-id="9a830-1149">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1149">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1150">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-1150">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a830-1151">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1151">Object</span></span>|<span data-ttu-id="9a830-1152">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1152">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1153">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1153">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a830-1154">function</span><span class="sxs-lookup"><span data-stu-id="9a830-1154">function</span></span>|<span data-ttu-id="9a830-1155">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1155">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1156">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-1156">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1157">Requirements</span></span>

|<span data-ttu-id="9a830-1158">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1158">Requirement</span></span>|<span data-ttu-id="9a830-1159">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1159">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1161">1,8</span><span class="sxs-lookup"><span data-stu-id="9a830-1161">1.8</span></span>|
|[<span data-ttu-id="9a830-1162">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1162">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1163">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1163">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1164">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1165">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-1165">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a830-1166">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9a830-1166">Returns:</span></span>

<span data-ttu-id="9a830-1167">Tipo: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="9a830-1167">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

##### <a name="example"></a><span data-ttu-id="9a830-1168">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1168">Example</span></span>

<span data-ttu-id="9a830-1169">O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="9a830-1169">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a><span data-ttu-id="9a830-1170">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="9a830-1170">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="9a830-1171">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="9a830-1171">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-1172">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="9a830-1172">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-1173">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1173">Requirements</span></span>

|<span data-ttu-id="9a830-1174">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1174">Requirement</span></span>|<span data-ttu-id="9a830-1175">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1175">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1176">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1177">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-1177">1.0</span></span>|
|[<span data-ttu-id="9a830-1178">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1178">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1179">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1181">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-1181">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a830-1182">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9a830-1182">Returns:</span></span>

<span data-ttu-id="9a830-1183">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-1183">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="9a830-1184">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1184">Example</span></span>

<span data-ttu-id="9a830-1185">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="9a830-1185">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a><span data-ttu-id="9a830-1186">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span><span class="sxs-lookup"><span data-stu-id="9a830-1186">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span></span>

<span data-ttu-id="9a830-1187">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="9a830-1187">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-1188">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="9a830-1188">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1189">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1189">Parameters</span></span>

|<span data-ttu-id="9a830-1190">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1190">Name</span></span>|<span data-ttu-id="9a830-1191">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1191">Type</span></span>|<span data-ttu-id="9a830-1192">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1192">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="9a830-1193">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="9a830-1193">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.8)|<span data-ttu-id="9a830-1194">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="9a830-1194">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1195">Requirements</span></span>

|<span data-ttu-id="9a830-1196">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1196">Requirement</span></span>|<span data-ttu-id="9a830-1197">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1197">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1199">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-1199">1.0</span></span>|
|[<span data-ttu-id="9a830-1200">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1201">Restrito</span><span class="sxs-lookup"><span data-stu-id="9a830-1201">Restricted</span></span>|
|[<span data-ttu-id="9a830-1202">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1203">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-1203">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a830-1204">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9a830-1204">Returns:</span></span>

<span data-ttu-id="9a830-1205">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="9a830-1205">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="9a830-1206">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="9a830-1206">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="9a830-1207">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="9a830-1207">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="9a830-1208">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-1208">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="9a830-1209">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="9a830-1209">Value of `entityType`</span></span>|<span data-ttu-id="9a830-1210">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="9a830-1210">Type of objects in returned array</span></span>|<span data-ttu-id="9a830-1211">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="9a830-1211">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="9a830-1212">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-1212">String</span></span>|<span data-ttu-id="9a830-1213">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="9a830-1213">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="9a830-1214">Contato</span><span class="sxs-lookup"><span data-stu-id="9a830-1214">Contact</span></span>|<span data-ttu-id="9a830-1215">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9a830-1215">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="9a830-1216">String</span><span class="sxs-lookup"><span data-stu-id="9a830-1216">String</span></span>|<span data-ttu-id="9a830-1217">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9a830-1217">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="9a830-1218">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="9a830-1218">MeetingSuggestion</span></span>|<span data-ttu-id="9a830-1219">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9a830-1219">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="9a830-1220">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="9a830-1220">PhoneNumber</span></span>|<span data-ttu-id="9a830-1221">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="9a830-1221">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="9a830-1222">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="9a830-1222">TaskSuggestion</span></span>|<span data-ttu-id="9a830-1223">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9a830-1223">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="9a830-1224">String</span><span class="sxs-lookup"><span data-stu-id="9a830-1224">String</span></span>|<span data-ttu-id="9a830-1225">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="9a830-1225">**Restricted**</span></span>|

<span data-ttu-id="9a830-1226">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span><span class="sxs-lookup"><span data-stu-id="9a830-1226">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span></span>

##### <a name="example"></a><span data-ttu-id="9a830-1227">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1227">Example</span></span>

<span data-ttu-id="9a830-1228">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="9a830-1228">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a><span data-ttu-id="9a830-1229">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span><span class="sxs-lookup"><span data-stu-id="9a830-1229">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span></span>

<span data-ttu-id="9a830-1230">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="9a830-1230">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-1231">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="9a830-1231">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9a830-1232">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="9a830-1232">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1233">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1233">Parameters</span></span>

|<span data-ttu-id="9a830-1234">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1234">Name</span></span>|<span data-ttu-id="9a830-1235">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1235">Type</span></span>|<span data-ttu-id="9a830-1236">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1236">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="9a830-1237">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9a830-1237">String</span></span>|<span data-ttu-id="9a830-1238">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="9a830-1238">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1239">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1239">Requirements</span></span>

|<span data-ttu-id="9a830-1240">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1240">Requirement</span></span>|<span data-ttu-id="9a830-1241">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1241">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1242">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1243">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-1243">1.0</span></span>|
|[<span data-ttu-id="9a830-1244">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1245">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1245">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1246">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1247">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-1247">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a830-1248">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9a830-1248">Returns:</span></span>

<span data-ttu-id="9a830-p174">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="9a830-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="9a830-1251">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span><span class="sxs-lookup"><span data-stu-id="9a830-1251">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span></span>

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="9a830-1252">getItemIdAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="9a830-1252">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="9a830-1253">Obtém de forma assíncrona a ID de um item salvo.</span><span class="sxs-lookup"><span data-stu-id="9a830-1253">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="9a830-1254">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="9a830-1254">Compose mode only.</span></span>

<span data-ttu-id="9a830-1255">Quando invocado, este método retorna a ID do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1255">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-1256">Se seu suplemento chamar `getItemIdAsync` um item no modo de redação (por exemplo, para `itemId` usar com o EWS ou a API REST), lembre-se de que, quando o Outlook estiver no modo cache, pode levar algum tempo para que o item seja sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="9a830-1256">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="9a830-1257">Até que o item seja sincronizado, `itemId` o não é reconhecido e usado retorna um erro.</span><span class="sxs-lookup"><span data-stu-id="9a830-1257">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1258">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1258">Parameters</span></span>

|<span data-ttu-id="9a830-1259">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1259">Name</span></span>|<span data-ttu-id="9a830-1260">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1260">Type</span></span>|<span data-ttu-id="9a830-1261">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-1261">Attributes</span></span>|<span data-ttu-id="9a830-1262">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1262">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="9a830-1263">Object</span><span class="sxs-lookup"><span data-stu-id="9a830-1263">Object</span></span>|<span data-ttu-id="9a830-1264">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1264">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1265">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-1265">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a830-1266">Object</span><span class="sxs-lookup"><span data-stu-id="9a830-1266">Object</span></span>|<span data-ttu-id="9a830-1267">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1268">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1268">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a830-1269">function</span><span class="sxs-lookup"><span data-stu-id="9a830-1269">function</span></span>||<span data-ttu-id="9a830-1270">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-1270">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9a830-1271">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9a830-1271">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9a830-1272">Erros</span><span class="sxs-lookup"><span data-stu-id="9a830-1272">Errors</span></span>

|<span data-ttu-id="9a830-1273">Código de erro</span><span class="sxs-lookup"><span data-stu-id="9a830-1273">Error code</span></span>|<span data-ttu-id="9a830-1274">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1274">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="9a830-1275">A ID não pode ser recuperada até que o item seja salvo.</span><span class="sxs-lookup"><span data-stu-id="9a830-1275">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1276">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1276">Requirements</span></span>

|<span data-ttu-id="9a830-1277">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1277">Requirement</span></span>|<span data-ttu-id="9a830-1278">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1278">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1279">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1280">1,8</span><span class="sxs-lookup"><span data-stu-id="9a830-1280">1.8</span></span>|
|[<span data-ttu-id="9a830-1281">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1281">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1282">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1283">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1283">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1284">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-1284">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a830-1285">Exemplos</span><span class="sxs-lookup"><span data-stu-id="9a830-1285">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="9a830-1286">O exemplo a seguir mostra a estrutura do `result` parâmetro que é passado para a função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1286">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="9a830-1287">A `value` propriedade contém a ID do item.</span><span class="sxs-lookup"><span data-stu-id="9a830-1287">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="9a830-1288">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="9a830-1288">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="9a830-1289">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="9a830-1289">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-1290">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="9a830-1290">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9a830-p178">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="9a830-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="9a830-1294">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="9a830-1294">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="9a830-1295">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="9a830-1295">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="9a830-p179">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="9a830-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-1299">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1299">Requirements</span></span>

|<span data-ttu-id="9a830-1300">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1300">Requirement</span></span>|<span data-ttu-id="9a830-1301">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1301">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1302">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1303">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-1303">1.0</span></span>|
|[<span data-ttu-id="9a830-1304">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1305">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1306">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1307">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-1307">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a830-1308">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9a830-1308">Returns:</span></span>

<span data-ttu-id="9a830-p180">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="9a830-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="9a830-1311">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="9a830-1311">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9a830-1312">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1312">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9a830-1313">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1313">Example</span></span>

<span data-ttu-id="9a830-1314">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="9a830-1314">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="9a830-1315">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="9a830-1315">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="9a830-1316">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="9a830-1316">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-1317">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="9a830-1317">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9a830-1318">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="9a830-1318">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="9a830-p181">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="9a830-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1321">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1321">Parameters</span></span>

|<span data-ttu-id="9a830-1322">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1322">Name</span></span>|<span data-ttu-id="9a830-1323">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1323">Type</span></span>|<span data-ttu-id="9a830-1324">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1324">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="9a830-1325">String</span><span class="sxs-lookup"><span data-stu-id="9a830-1325">String</span></span>|<span data-ttu-id="9a830-1326">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="9a830-1326">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1327">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1327">Requirements</span></span>

|<span data-ttu-id="9a830-1328">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1328">Requirement</span></span>|<span data-ttu-id="9a830-1329">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1329">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1330">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1330">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1331">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-1331">1.0</span></span>|
|[<span data-ttu-id="9a830-1332">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1332">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1333">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1333">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1334">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1334">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1335">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-1335">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a830-1336">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9a830-1336">Returns:</span></span>

<span data-ttu-id="9a830-1337">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="9a830-1337">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="9a830-1338">Tipo: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="9a830-1338">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="9a830-1339">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1339">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="9a830-1340">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="9a830-1340">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="9a830-1341">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a830-1341">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="9a830-1342">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retornará uma cadeia de caracteres vazia para os dados selecionados.</span><span class="sxs-lookup"><span data-stu-id="9a830-1342">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.</span></span> <span data-ttu-id="9a830-1343">Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="9a830-1343">If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-1344">No Outlook na Web, o método retorna a cadeia de caracteres “null” se nenhum texto for selecionado, mas o cursor estiver no corpo.</span><span class="sxs-lookup"><span data-stu-id="9a830-1344">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="9a830-1345">Para verificar essa situação, confira o exemplo mais adiante nesta seção.</span><span class="sxs-lookup"><span data-stu-id="9a830-1345">To check for this situation, see the example later in this section.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1346">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1346">Parameters</span></span>

|<span data-ttu-id="9a830-1347">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1347">Name</span></span>|<span data-ttu-id="9a830-1348">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1348">Type</span></span>|<span data-ttu-id="9a830-1349">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-1349">Attributes</span></span>|<span data-ttu-id="9a830-1350">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1350">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="9a830-1351">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="9a830-1351">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="9a830-p184">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="9a830-p184">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="9a830-1355">Object</span><span class="sxs-lookup"><span data-stu-id="9a830-1355">Object</span></span>|<span data-ttu-id="9a830-1356">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1356">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1357">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-1357">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a830-1358">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1358">Object</span></span>|<span data-ttu-id="9a830-1359">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1359">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1360">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1360">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a830-1361">function</span><span class="sxs-lookup"><span data-stu-id="9a830-1361">function</span></span>||<span data-ttu-id="9a830-1362">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-1362">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9a830-1363">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="9a830-1363">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="9a830-1364">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="9a830-1364">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1365">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1365">Requirements</span></span>

|<span data-ttu-id="9a830-1366">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1366">Requirement</span></span>|<span data-ttu-id="9a830-1367">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1367">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1368">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1369">1.2</span><span class="sxs-lookup"><span data-stu-id="9a830-1369">1.2</span></span>|
|[<span data-ttu-id="9a830-1370">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1371">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1371">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1372">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1373">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-1373">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a830-1374">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9a830-1374">Returns:</span></span>

<span data-ttu-id="9a830-1375">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="9a830-1375">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="9a830-1376">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="9a830-1376">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="9a830-1377">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1377">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  // Handle where Outlook on the web erroneously returns "null" instead of empty string.
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookWebApp'
      && asyncResult.value.endPosition === asyncResult.value.startPosition) {
    text = "";
  }

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a><span data-ttu-id="9a830-1378">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="9a830-1378">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="9a830-1379">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="9a830-1379">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="9a830-1380">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="9a830-1380">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-1381">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="9a830-1381">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-1382">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1382">Requirements</span></span>

|<span data-ttu-id="9a830-1383">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1383">Requirement</span></span>|<span data-ttu-id="9a830-1384">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1384">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1385">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1385">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1386">1.6</span><span class="sxs-lookup"><span data-stu-id="9a830-1386">1.6</span></span>|
|[<span data-ttu-id="9a830-1387">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1387">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1388">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1388">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1389">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1389">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1390">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-1390">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a830-1391">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9a830-1391">Returns:</span></span>

<span data-ttu-id="9a830-1392">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="9a830-1392">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="9a830-1393">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1393">Example</span></span>

<span data-ttu-id="9a830-1394">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="9a830-1394">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="9a830-1395">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="9a830-1395">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="9a830-p187">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="9a830-p187">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-1398">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="9a830-1398">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9a830-p188">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="9a830-p188">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="9a830-1402">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="9a830-1402">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="9a830-1403">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="9a830-1403">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="9a830-p189">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="9a830-p189">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a830-1407">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1407">Requirements</span></span>

|<span data-ttu-id="9a830-1408">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1408">Requirement</span></span>|<span data-ttu-id="9a830-1409">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1409">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1410">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1411">1.6</span><span class="sxs-lookup"><span data-stu-id="9a830-1411">1.6</span></span>|
|[<span data-ttu-id="9a830-1412">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1413">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1414">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1415">Read</span><span class="sxs-lookup"><span data-stu-id="9a830-1415">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a830-1416">Retorna:</span><span class="sxs-lookup"><span data-stu-id="9a830-1416">Returns:</span></span>

<span data-ttu-id="9a830-p190">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="9a830-p190">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="9a830-1419">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1419">Example</span></span>

<span data-ttu-id="9a830-1420">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="9a830-1420">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="9a830-1421">getSharedPropertiesAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="9a830-1421">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="9a830-1422">Obtém as propriedades do compromisso ou da mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="9a830-1422">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1423">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1423">Parameters</span></span>

|<span data-ttu-id="9a830-1424">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1424">Name</span></span>|<span data-ttu-id="9a830-1425">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1425">Type</span></span>|<span data-ttu-id="9a830-1426">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-1426">Attributes</span></span>|<span data-ttu-id="9a830-1427">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1427">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="9a830-1428">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1428">Object</span></span>|<span data-ttu-id="9a830-1429">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1429">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1430">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-1430">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a830-1431">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1431">Object</span></span>|<span data-ttu-id="9a830-1432">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1432">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1433">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1433">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a830-1434">function</span><span class="sxs-lookup"><span data-stu-id="9a830-1434">function</span></span>||<span data-ttu-id="9a830-1435">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-1435">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9a830-1436">As propriedades compartilhadas são fornecidas [`SharedProperties`](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) como um objeto `asyncResult.value` na propriedade.</span><span class="sxs-lookup"><span data-stu-id="9a830-1436">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="9a830-1437">Este objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="9a830-1437">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1438">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1438">Requirements</span></span>

|<span data-ttu-id="9a830-1439">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1439">Requirement</span></span>|<span data-ttu-id="9a830-1440">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1440">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1441">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1442">1,8</span><span class="sxs-lookup"><span data-stu-id="9a830-1442">1.8</span></span>|
|[<span data-ttu-id="9a830-1443">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1444">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1445">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-1445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1446">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-1446">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-1447">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1447">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="9a830-1448">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9a830-1448">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="9a830-1449">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="9a830-1449">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="9a830-p192">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="9a830-p192">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1453">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1453">Parameters</span></span>

|<span data-ttu-id="9a830-1454">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1454">Name</span></span>|<span data-ttu-id="9a830-1455">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1455">Type</span></span>|<span data-ttu-id="9a830-1456">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-1456">Attributes</span></span>|<span data-ttu-id="9a830-1457">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1457">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="9a830-1458">function</span><span class="sxs-lookup"><span data-stu-id="9a830-1458">function</span></span>||<span data-ttu-id="9a830-1459">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-1459">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9a830-1460">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.8) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9a830-1460">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.8) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="9a830-1461">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="9a830-1461">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="9a830-1462">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1462">Object</span></span>|<span data-ttu-id="9a830-1463">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1463">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1464">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1464">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="9a830-1465">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1465">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1466">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1466">Requirements</span></span>

|<span data-ttu-id="9a830-1467">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1467">Requirement</span></span>|<span data-ttu-id="9a830-1468">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1468">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1469">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1470">1.0</span><span class="sxs-lookup"><span data-stu-id="9a830-1470">1.0</span></span>|
|[<span data-ttu-id="9a830-1471">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1472">ReadItem</span></span>|
|[<span data-ttu-id="9a830-1473">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-1473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1474">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-1474">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-1475">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1475">Example</span></span>

<span data-ttu-id="9a830-p195">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="9a830-p195">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="9a830-1479">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a830-1479">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="9a830-1480">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a830-1480">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="9a830-1481">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="9a830-1481">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="9a830-1482">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="9a830-1482">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="9a830-1483">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="9a830-1483">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="9a830-1484">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1484">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1485">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1485">Parameters</span></span>

|<span data-ttu-id="9a830-1486">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1486">Name</span></span>|<span data-ttu-id="9a830-1487">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1487">Type</span></span>|<span data-ttu-id="9a830-1488">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-1488">Attributes</span></span>|<span data-ttu-id="9a830-1489">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1489">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="9a830-1490">String</span><span class="sxs-lookup"><span data-stu-id="9a830-1490">String</span></span>||<span data-ttu-id="9a830-1491">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="9a830-1491">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="9a830-1492">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1492">Object</span></span>|<span data-ttu-id="9a830-1493">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1493">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1494">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-1494">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a830-1495">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1495">Object</span></span>|<span data-ttu-id="9a830-1496">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1496">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1497">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1497">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a830-1498">function</span><span class="sxs-lookup"><span data-stu-id="9a830-1498">function</span></span>|<span data-ttu-id="9a830-1499">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1499">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1500">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-1500">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9a830-1501">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="9a830-1501">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9a830-1502">Erros</span><span class="sxs-lookup"><span data-stu-id="9a830-1502">Errors</span></span>

|<span data-ttu-id="9a830-1503">Código de erro</span><span class="sxs-lookup"><span data-stu-id="9a830-1503">Error code</span></span>|<span data-ttu-id="9a830-1504">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1504">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="9a830-1505">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="9a830-1505">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1506">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1506">Requirements</span></span>

|<span data-ttu-id="9a830-1507">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1507">Requirement</span></span>|<span data-ttu-id="9a830-1508">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1508">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1509">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1510">1.1</span><span class="sxs-lookup"><span data-stu-id="9a830-1510">1.1</span></span>|
|[<span data-ttu-id="9a830-1511">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1512">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1512">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a830-1513">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1514">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-1514">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-1515">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1515">Example</span></span>

<span data-ttu-id="9a830-1516">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="9a830-1516">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="9a830-1517">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a830-1517">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="9a830-1518">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="9a830-1518">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="9a830-1519">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="9a830-1519">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1520">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1520">Parameters</span></span>

| <span data-ttu-id="9a830-1521">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1521">Name</span></span> | <span data-ttu-id="9a830-1522">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1522">Type</span></span> | <span data-ttu-id="9a830-1523">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-1523">Attributes</span></span> | <span data-ttu-id="9a830-1524">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1524">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="9a830-1525">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="9a830-1525">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="9a830-1526">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="9a830-1526">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="9a830-1527">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1527">Object</span></span> | <span data-ttu-id="9a830-1528">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1528">&lt;optional&gt;</span></span> | <span data-ttu-id="9a830-1529">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-1529">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9a830-1530">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1530">Object</span></span> | <span data-ttu-id="9a830-1531">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1531">&lt;optional&gt;</span></span> | <span data-ttu-id="9a830-1532">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1532">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="9a830-1533">function</span><span class="sxs-lookup"><span data-stu-id="9a830-1533">function</span></span>| <span data-ttu-id="9a830-1534">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1534">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1535">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-1535">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1536">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1536">Requirements</span></span>

|<span data-ttu-id="9a830-1537">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1537">Requirement</span></span>| <span data-ttu-id="9a830-1538">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1538">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1539">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a830-1540">1.7</span><span class="sxs-lookup"><span data-stu-id="9a830-1540">1.7</span></span> |
|[<span data-ttu-id="9a830-1541">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a830-1542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1542">ReadItem</span></span> |
|[<span data-ttu-id="9a830-1543">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a830-1543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9a830-1544">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9a830-1544">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="9a830-1545">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="9a830-1545">saveAsync([options], callback)</span></span>

<span data-ttu-id="9a830-1546">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="9a830-1546">Asynchronously saves an item.</span></span>

<span data-ttu-id="9a830-1547">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1547">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="9a830-1548">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="9a830-1548">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="9a830-1549">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="9a830-1549">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-1550">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="9a830-1550">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="9a830-1551">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="9a830-1551">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="9a830-p199">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="9a830-p199">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="9a830-1555">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="9a830-1555">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="9a830-1556">O Outlook no Mac não dá suporte ao salvamento de reuniões.</span><span class="sxs-lookup"><span data-stu-id="9a830-1556">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="9a830-1557">O método `saveAsync` falha quando chamado a partir de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="9a830-1557">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="9a830-1558">Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="9a830-1558">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="9a830-1559">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="9a830-1559">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1560">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1560">Parameters</span></span>

|<span data-ttu-id="9a830-1561">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1561">Name</span></span>|<span data-ttu-id="9a830-1562">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1562">Type</span></span>|<span data-ttu-id="9a830-1563">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-1563">Attributes</span></span>|<span data-ttu-id="9a830-1564">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1564">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="9a830-1565">Object</span><span class="sxs-lookup"><span data-stu-id="9a830-1565">Object</span></span>|<span data-ttu-id="9a830-1566">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1566">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1567">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-1567">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a830-1568">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1568">Object</span></span>|<span data-ttu-id="9a830-1569">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1569">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1570">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1570">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a830-1571">function</span><span class="sxs-lookup"><span data-stu-id="9a830-1571">function</span></span>||<span data-ttu-id="9a830-1572">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-1572">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9a830-1573">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9a830-1573">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1574">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1574">Requirements</span></span>

|<span data-ttu-id="9a830-1575">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1575">Requirement</span></span>|<span data-ttu-id="9a830-1576">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1576">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1577">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1577">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1578">1.3</span><span class="sxs-lookup"><span data-stu-id="9a830-1578">1.3</span></span>|
|[<span data-ttu-id="9a830-1579">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1579">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1580">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1580">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a830-1581">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1581">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1582">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-1582">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a830-1583">Exemplos</span><span class="sxs-lookup"><span data-stu-id="9a830-1583">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="9a830-p201">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="9a830-p201">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="9a830-1586">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="9a830-1586">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="9a830-1587">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="9a830-1587">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="9a830-p202">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="9a830-p202">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a830-1591">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="9a830-1591">Parameters</span></span>

|<span data-ttu-id="9a830-1592">Nome</span><span class="sxs-lookup"><span data-stu-id="9a830-1592">Name</span></span>|<span data-ttu-id="9a830-1593">Tipo</span><span class="sxs-lookup"><span data-stu-id="9a830-1593">Type</span></span>|<span data-ttu-id="9a830-1594">Atributos</span><span class="sxs-lookup"><span data-stu-id="9a830-1594">Attributes</span></span>|<span data-ttu-id="9a830-1595">Descrição</span><span class="sxs-lookup"><span data-stu-id="9a830-1595">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="9a830-1596">String</span><span class="sxs-lookup"><span data-stu-id="9a830-1596">String</span></span>||<span data-ttu-id="9a830-p203">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="9a830-p203">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="9a830-1600">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1600">Object</span></span>|<span data-ttu-id="9a830-1601">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1601">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1602">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a830-1602">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a830-1603">Objeto</span><span class="sxs-lookup"><span data-stu-id="9a830-1603">Object</span></span>|<span data-ttu-id="9a830-1604">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1604">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1605">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="9a830-1605">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="9a830-1606">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="9a830-1606">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="9a830-1607">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a830-1607">&lt;optional&gt;</span></span>|<span data-ttu-id="9a830-1608">Se `text`, o estilo atual é aplicado nos clientes do Outlook na Web e do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="9a830-1608">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="9a830-1609">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="9a830-1609">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="9a830-1610">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook na Web e o estilo padrão é aplicado nos clientes do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="9a830-1610">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="9a830-1611">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="9a830-1611">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="9a830-1612">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="9a830-1612">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="9a830-1613">function</span><span class="sxs-lookup"><span data-stu-id="9a830-1613">function</span></span>||<span data-ttu-id="9a830-1614">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a830-1614">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a830-1615">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9a830-1615">Requirements</span></span>

|<span data-ttu-id="9a830-1616">Requisito</span><span class="sxs-lookup"><span data-stu-id="9a830-1616">Requirement</span></span>|<span data-ttu-id="9a830-1617">Valor</span><span class="sxs-lookup"><span data-stu-id="9a830-1617">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a830-1618">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9a830-1618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a830-1619">1.2</span><span class="sxs-lookup"><span data-stu-id="9a830-1619">1.2</span></span>|
|[<span data-ttu-id="9a830-1620">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9a830-1620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a830-1621">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a830-1621">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a830-1622">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9a830-1622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a830-1623">Escrever</span><span class="sxs-lookup"><span data-stu-id="9a830-1623">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9a830-1624">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9a830-1624">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
