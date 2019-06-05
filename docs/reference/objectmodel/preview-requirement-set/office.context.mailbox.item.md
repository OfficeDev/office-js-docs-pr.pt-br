---
title: Office. Context. Mailbox. Item-visualização do conjunto de requisitos
description: ''
ms.date: 06/03/2019
localization_priority: Normal
ms.openlocfilehash: 3dad9133fb23f6190e58eab94dc1724c18ac9d40
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706355"
---
# <a name="item"></a><span data-ttu-id="44fc2-102">item</span><span class="sxs-lookup"><span data-stu-id="44fc2-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="44fc2-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="44fc2-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="44fc2-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="44fc2-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-106">Requirements</span></span>

|<span data-ttu-id="44fc2-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-107">Requirement</span></span>|<span data-ttu-id="44fc2-108">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-110">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-110">1.0</span></span>|
|[<span data-ttu-id="44fc2-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="44fc2-112">Restricted</span></span>|
|[<span data-ttu-id="44fc2-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="44fc2-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="44fc2-115">Members and methods</span></span>

| <span data-ttu-id="44fc2-116">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-116">Member</span></span> | <span data-ttu-id="44fc2-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="44fc2-118">attachments</span><span class="sxs-lookup"><span data-stu-id="44fc2-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="44fc2-119">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-119">Member</span></span> |
| [<span data-ttu-id="44fc2-120">bcc</span><span class="sxs-lookup"><span data-stu-id="44fc2-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="44fc2-121">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-121">Member</span></span> |
| [<span data-ttu-id="44fc2-122">body</span><span class="sxs-lookup"><span data-stu-id="44fc2-122">body</span></span>](#body-body) | <span data-ttu-id="44fc2-123">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-123">Member</span></span> |
| [<span data-ttu-id="44fc2-124">Categorias</span><span class="sxs-lookup"><span data-stu-id="44fc2-124">categories</span></span>](#categories-categories) | <span data-ttu-id="44fc2-125">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-125">Member</span></span> |
| [<span data-ttu-id="44fc2-126">cc</span><span class="sxs-lookup"><span data-stu-id="44fc2-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="44fc2-127">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-127">Member</span></span> |
| [<span data-ttu-id="44fc2-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="44fc2-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="44fc2-129">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-129">Member</span></span> |
| [<span data-ttu-id="44fc2-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="44fc2-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="44fc2-131">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-131">Member</span></span> |
| [<span data-ttu-id="44fc2-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="44fc2-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="44fc2-133">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-133">Member</span></span> |
| [<span data-ttu-id="44fc2-134">end</span><span class="sxs-lookup"><span data-stu-id="44fc2-134">end</span></span>](#end-datetime) | <span data-ttu-id="44fc2-135">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-135">Member</span></span> |
| [<span data-ttu-id="44fc2-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="44fc2-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="44fc2-137">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-137">Member</span></span> |
| [<span data-ttu-id="44fc2-138">from</span><span class="sxs-lookup"><span data-stu-id="44fc2-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="44fc2-139">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-139">Member</span></span> |
| [<span data-ttu-id="44fc2-140">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="44fc2-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="44fc2-141">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-141">Member</span></span> |
| [<span data-ttu-id="44fc2-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="44fc2-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="44fc2-143">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-143">Member</span></span> |
| [<span data-ttu-id="44fc2-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="44fc2-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="44fc2-145">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-145">Member</span></span> |
| [<span data-ttu-id="44fc2-146">itemId</span><span class="sxs-lookup"><span data-stu-id="44fc2-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="44fc2-147">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-147">Member</span></span> |
| [<span data-ttu-id="44fc2-148">itemType</span><span class="sxs-lookup"><span data-stu-id="44fc2-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="44fc2-149">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-149">Member</span></span> |
| [<span data-ttu-id="44fc2-150">location</span><span class="sxs-lookup"><span data-stu-id="44fc2-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="44fc2-151">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-151">Member</span></span> |
| [<span data-ttu-id="44fc2-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="44fc2-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="44fc2-153">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-153">Member</span></span> |
| [<span data-ttu-id="44fc2-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="44fc2-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="44fc2-155">Member</span><span class="sxs-lookup"><span data-stu-id="44fc2-155">Member</span></span> |
| [<span data-ttu-id="44fc2-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="44fc2-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="44fc2-157">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-157">Member</span></span> |
| [<span data-ttu-id="44fc2-158">organizer</span><span class="sxs-lookup"><span data-stu-id="44fc2-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="44fc2-159">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-159">Member</span></span> |
| [<span data-ttu-id="44fc2-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="44fc2-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="44fc2-161">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-161">Member</span></span> |
| [<span data-ttu-id="44fc2-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="44fc2-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="44fc2-163">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-163">Member</span></span> |
| [<span data-ttu-id="44fc2-164">sender</span><span class="sxs-lookup"><span data-stu-id="44fc2-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="44fc2-165">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-165">Member</span></span> |
| [<span data-ttu-id="44fc2-166">seriesid</span><span class="sxs-lookup"><span data-stu-id="44fc2-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="44fc2-167">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-167">Member</span></span> |
| [<span data-ttu-id="44fc2-168">start</span><span class="sxs-lookup"><span data-stu-id="44fc2-168">start</span></span>](#start-datetime) | <span data-ttu-id="44fc2-169">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-169">Member</span></span> |
| [<span data-ttu-id="44fc2-170">subject</span><span class="sxs-lookup"><span data-stu-id="44fc2-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="44fc2-171">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-171">Member</span></span> |
| [<span data-ttu-id="44fc2-172">to</span><span class="sxs-lookup"><span data-stu-id="44fc2-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="44fc2-173">Membro</span><span class="sxs-lookup"><span data-stu-id="44fc2-173">Member</span></span> |
| [<span data-ttu-id="44fc2-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="44fc2-175">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-175">Method</span></span> |
| [<span data-ttu-id="44fc2-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="44fc2-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="44fc2-177">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-177">Method</span></span> |
| [<span data-ttu-id="44fc2-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="44fc2-179">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-179">Method</span></span> |
| [<span data-ttu-id="44fc2-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="44fc2-181">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-181">Method</span></span> |
| [<span data-ttu-id="44fc2-182">close</span><span class="sxs-lookup"><span data-stu-id="44fc2-182">close</span></span>](#close) | <span data-ttu-id="44fc2-183">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-183">Method</span></span> |
| [<span data-ttu-id="44fc2-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="44fc2-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="44fc2-185">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-185">Method</span></span> |
| [<span data-ttu-id="44fc2-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="44fc2-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="44fc2-187">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-187">Method</span></span> |
| [<span data-ttu-id="44fc2-188">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="44fc2-189">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-189">Method</span></span> |
| [<span data-ttu-id="44fc2-190">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="44fc2-191">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-191">Method</span></span> |
| [<span data-ttu-id="44fc2-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="44fc2-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="44fc2-193">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-193">Method</span></span> |
| [<span data-ttu-id="44fc2-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="44fc2-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="44fc2-195">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-195">Method</span></span> |
| [<span data-ttu-id="44fc2-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="44fc2-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="44fc2-197">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-197">Method</span></span> |
| [<span data-ttu-id="44fc2-198">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="44fc2-199">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-199">Method</span></span> |
| [<span data-ttu-id="44fc2-200">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="44fc2-201">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-201">Method</span></span> |
| [<span data-ttu-id="44fc2-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="44fc2-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="44fc2-203">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-203">Method</span></span> |
| [<span data-ttu-id="44fc2-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="44fc2-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="44fc2-205">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-205">Method</span></span> |
| [<span data-ttu-id="44fc2-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="44fc2-207">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-207">Method</span></span> |
| [<span data-ttu-id="44fc2-208">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="44fc2-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="44fc2-209">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-209">Method</span></span> |
| [<span data-ttu-id="44fc2-210">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="44fc2-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="44fc2-211">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-211">Method</span></span> |
| [<span data-ttu-id="44fc2-212">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="44fc2-213">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-213">Method</span></span> |
| [<span data-ttu-id="44fc2-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="44fc2-215">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-215">Method</span></span> |
| [<span data-ttu-id="44fc2-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="44fc2-217">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-217">Method</span></span> |
| [<span data-ttu-id="44fc2-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="44fc2-219">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-219">Method</span></span> |
| [<span data-ttu-id="44fc2-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="44fc2-221">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-221">Method</span></span> |
| [<span data-ttu-id="44fc2-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="44fc2-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="44fc2-223">Método</span><span class="sxs-lookup"><span data-stu-id="44fc2-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="44fc2-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-224">Example</span></span>

<span data-ttu-id="44fc2-225">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="44fc2-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="44fc2-226">Membros</span><span class="sxs-lookup"><span data-stu-id="44fc2-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="44fc2-227">anexos: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="44fc2-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="44fc2-228">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="44fc2-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="44fc2-229">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="44fc2-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-230">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="44fc2-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="44fc2-231">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="44fc2-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-232">Type</span></span>

*   <span data-ttu-id="44fc2-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="44fc2-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-234">Requirements</span></span>

|<span data-ttu-id="44fc2-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-235">Requirement</span></span>|<span data-ttu-id="44fc2-236">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-238">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-238">1.0</span></span>|
|[<span data-ttu-id="44fc2-239">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-240">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-241">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-242">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-243">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-243">Example</span></span>

<span data-ttu-id="44fc2-244">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="44fc2-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="44fc2-245">CCO: [destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44fc2-245">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="44fc2-246">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="44fc2-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="44fc2-247">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="44fc2-247">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-248">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-248">Type</span></span>

*   [<span data-ttu-id="44fc2-249">Destinatários</span><span class="sxs-lookup"><span data-stu-id="44fc2-249">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="44fc2-250">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-250">Requirements</span></span>

|<span data-ttu-id="44fc2-251">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-251">Requirement</span></span>|<span data-ttu-id="44fc2-252">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-253">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-254">1.1</span><span class="sxs-lookup"><span data-stu-id="44fc2-254">1.1</span></span>|
|[<span data-ttu-id="44fc2-255">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-256">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-257">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-258">Escrever</span><span class="sxs-lookup"><span data-stu-id="44fc2-258">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-259">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-259">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="44fc2-260">corpo: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="44fc2-260">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="44fc2-261">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-261">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-262">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-262">Type</span></span>

*   [<span data-ttu-id="44fc2-263">Body</span><span class="sxs-lookup"><span data-stu-id="44fc2-263">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="44fc2-264">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-264">Requirements</span></span>

|<span data-ttu-id="44fc2-265">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-265">Requirement</span></span>|<span data-ttu-id="44fc2-266">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-267">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-268">1.1</span><span class="sxs-lookup"><span data-stu-id="44fc2-268">1.1</span></span>|
|[<span data-ttu-id="44fc2-269">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-270">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-271">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-272">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-273">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-273">Example</span></span>

<span data-ttu-id="44fc2-274">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="44fc2-274">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="44fc2-275">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-275">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="44fc2-276">Categorias: [categorias](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="44fc2-276">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="44fc2-277">Obtém um objeto que fornece métodos para gerenciar as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-277">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-278">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="44fc2-278">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-279">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-279">Type</span></span>

*   [<span data-ttu-id="44fc2-280">Categories</span><span class="sxs-lookup"><span data-stu-id="44fc2-280">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="44fc2-281">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-281">Requirements</span></span>

|<span data-ttu-id="44fc2-282">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-282">Requirement</span></span>|<span data-ttu-id="44fc2-283">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-284">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-285">Visualização</span><span class="sxs-lookup"><span data-stu-id="44fc2-285">Preview</span></span>|
|[<span data-ttu-id="44fc2-286">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-287">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-288">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-289">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-289">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-290">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-290">Example</span></span>

<span data-ttu-id="44fc2-291">Este exemplo obtém as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-291">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="44fc2-292">[destinatários](/javascript/api/outlook/office.recipients) [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|CC: Array. <</span><span class="sxs-lookup"><span data-stu-id="44fc2-292">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="44fc2-293">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="44fc2-293">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="44fc2-294">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="44fc2-294">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44fc2-295">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="44fc2-295">Read mode</span></span>

<span data-ttu-id="44fc2-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="44fc2-298">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="44fc2-298">Compose mode</span></span>

<span data-ttu-id="44fc2-299">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="44fc2-299">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="44fc2-300">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-300">Type</span></span>

*   <span data-ttu-id="44fc2-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44fc2-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-302">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-302">Requirements</span></span>

|<span data-ttu-id="44fc2-303">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-303">Requirement</span></span>|<span data-ttu-id="44fc2-304">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-305">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-306">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-306">1.0</span></span>|
|[<span data-ttu-id="44fc2-307">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-308">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-309">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-310">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-310">Compose or Read</span></span>|

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="44fc2-311">(Nullable) Conversation: String</span><span class="sxs-lookup"><span data-stu-id="44fc2-311">(nullable) conversationId: String</span></span>

<span data-ttu-id="44fc2-312">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="44fc2-312">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="44fc2-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="44fc2-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-317">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-317">Type</span></span>

*   <span data-ttu-id="44fc2-318">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-318">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-319">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-319">Requirements</span></span>

|<span data-ttu-id="44fc2-320">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-320">Requirement</span></span>|<span data-ttu-id="44fc2-321">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-321">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-322">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-322">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-323">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-323">1.0</span></span>|
|[<span data-ttu-id="44fc2-324">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-324">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-325">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-325">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-326">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-326">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-327">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-327">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-328">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-328">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="44fc2-329">dateTimeCreated: data</span><span class="sxs-lookup"><span data-stu-id="44fc2-329">dateTimeCreated: Date</span></span>

<span data-ttu-id="44fc2-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-332">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-332">Type</span></span>

*   <span data-ttu-id="44fc2-333">Data</span><span class="sxs-lookup"><span data-stu-id="44fc2-333">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-334">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-334">Requirements</span></span>

|<span data-ttu-id="44fc2-335">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-335">Requirement</span></span>|<span data-ttu-id="44fc2-336">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-337">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-338">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-338">1.0</span></span>|
|[<span data-ttu-id="44fc2-339">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-340">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-341">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-342">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-343">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-343">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="44fc2-344">dateTimeModified: data</span><span class="sxs-lookup"><span data-stu-id="44fc2-344">dateTimeModified: Date</span></span>

<span data-ttu-id="44fc2-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-347">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="44fc2-347">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-348">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-348">Type</span></span>

*   <span data-ttu-id="44fc2-349">Data</span><span class="sxs-lookup"><span data-stu-id="44fc2-349">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-350">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-350">Requirements</span></span>

|<span data-ttu-id="44fc2-351">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-351">Requirement</span></span>|<span data-ttu-id="44fc2-352">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-353">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-354">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-354">1.0</span></span>|
|[<span data-ttu-id="44fc2-355">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-356">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-357">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-358">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-358">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-359">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-359">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="44fc2-360">fim: data | [Tempo](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="44fc2-360">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="44fc2-361">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="44fc2-361">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="44fc2-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44fc2-364">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="44fc2-364">Read mode</span></span>

<span data-ttu-id="44fc2-365">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-365">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="44fc2-366">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="44fc2-366">Compose mode</span></span>

<span data-ttu-id="44fc2-367">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-367">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="44fc2-368">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="44fc2-368">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="44fc2-369">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-369">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="44fc2-370">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-370">Type</span></span>

*   <span data-ttu-id="44fc2-371">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="44fc2-371">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-372">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-372">Requirements</span></span>

|<span data-ttu-id="44fc2-373">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-373">Requirement</span></span>|<span data-ttu-id="44fc2-374">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-375">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-376">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-376">1.0</span></span>|
|[<span data-ttu-id="44fc2-377">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-378">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-379">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-380">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-380">Compose or Read</span></span>|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="44fc2-381">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="44fc2-381">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="44fc2-382">Obtém ou define os locais de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-382">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44fc2-383">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="44fc2-383">Read mode</span></span>

<span data-ttu-id="44fc2-384">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que permite que você obtenha o conjunto de locais (cada um representado por um objeto [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associado ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="44fc2-385">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="44fc2-385">Compose mode</span></span>

<span data-ttu-id="44fc2-386">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que fornece métodos para obter, remover ou adicionar locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-386">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-387">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-387">Type</span></span>

*   [<span data-ttu-id="44fc2-388">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="44fc2-388">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="44fc2-389">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-389">Requirements</span></span>

|<span data-ttu-id="44fc2-390">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-390">Requirement</span></span>|<span data-ttu-id="44fc2-391">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-391">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-392">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-392">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-393">Visualização</span><span class="sxs-lookup"><span data-stu-id="44fc2-393">Preview</span></span>|
|[<span data-ttu-id="44fc2-394">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-395">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-396">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-396">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-397">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-397">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-398">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-398">Example</span></span>

<span data-ttu-id="44fc2-399">O exemplo a seguir obtém os locais atuais associados ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-399">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="44fc2-400">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="44fc2-400">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="44fc2-401">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="44fc2-401">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="44fc2-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-404">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-404">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44fc2-405">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="44fc2-405">Read mode</span></span>

<span data-ttu-id="44fc2-406">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="44fc2-406">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="44fc2-407">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="44fc2-407">Compose mode</span></span>

<span data-ttu-id="44fc2-408">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="44fc2-408">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="44fc2-409">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-409">Type</span></span>

*   <span data-ttu-id="44fc2-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="44fc2-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-411">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-411">Requirements</span></span>

|<span data-ttu-id="44fc2-412">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-412">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="44fc2-413">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-414">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-414">1.0</span></span>|<span data-ttu-id="44fc2-415">1.7</span><span class="sxs-lookup"><span data-stu-id="44fc2-415">1.7</span></span>|
|[<span data-ttu-id="44fc2-416">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-417">ReadItem</span></span>|<span data-ttu-id="44fc2-418">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-418">ReadWriteItem</span></span>|
|[<span data-ttu-id="44fc2-419">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-419">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-420">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-420">Read</span></span>|<span data-ttu-id="44fc2-421">Escrever</span><span class="sxs-lookup"><span data-stu-id="44fc2-421">Compose</span></span>|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="44fc2-422">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="44fc2-422">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="44fc2-423">Obtém ou define os cabeçalhos de Internet de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="44fc2-423">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-424">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-424">Type</span></span>

*   [<span data-ttu-id="44fc2-425">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="44fc2-425">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="44fc2-426">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-426">Requirements</span></span>

|<span data-ttu-id="44fc2-427">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-427">Requirement</span></span>|<span data-ttu-id="44fc2-428">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-429">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-430">Visualização</span><span class="sxs-lookup"><span data-stu-id="44fc2-430">Preview</span></span>|
|[<span data-ttu-id="44fc2-431">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-432">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-433">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-434">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-435">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-435">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="44fc2-436">internetMessageId: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="44fc2-436">internetMessageId: String</span></span>

<span data-ttu-id="44fc2-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-439">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-439">Type</span></span>

*   <span data-ttu-id="44fc2-440">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-440">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-441">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-441">Requirements</span></span>

|<span data-ttu-id="44fc2-442">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-442">Requirement</span></span>|<span data-ttu-id="44fc2-443">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-444">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-445">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-445">1.0</span></span>|
|[<span data-ttu-id="44fc2-446">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-447">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-448">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-449">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-449">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-450">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-450">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="44fc2-451">doclass: String</span><span class="sxs-lookup"><span data-stu-id="44fc2-451">itemClass: String</span></span>

<span data-ttu-id="44fc2-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="44fc2-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="44fc2-456">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-456">Type</span></span>|<span data-ttu-id="44fc2-457">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-457">Description</span></span>|<span data-ttu-id="44fc2-458">classe de item</span><span class="sxs-lookup"><span data-stu-id="44fc2-458">item class</span></span>|
|---|---|---|
|<span data-ttu-id="44fc2-459">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="44fc2-459">Appointment items</span></span>|<span data-ttu-id="44fc2-460">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-460">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="44fc2-461">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="44fc2-461">Message items</span></span>|<span data-ttu-id="44fc2-462">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="44fc2-462">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="44fc2-463">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-463">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-464">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-464">Type</span></span>

*   <span data-ttu-id="44fc2-465">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-465">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-466">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-466">Requirements</span></span>

|<span data-ttu-id="44fc2-467">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-467">Requirement</span></span>|<span data-ttu-id="44fc2-468">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-469">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-470">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-470">1.0</span></span>|
|[<span data-ttu-id="44fc2-471">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-472">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-473">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-474">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-474">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-475">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-475">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="44fc2-476">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="44fc2-476">(nullable) itemId: String</span></span>

<span data-ttu-id="44fc2-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-479">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="44fc2-479">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="44fc2-480">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="44fc2-480">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="44fc2-481">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="44fc2-481">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="44fc2-482">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="44fc2-482">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="44fc2-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-485">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-485">Type</span></span>

*   <span data-ttu-id="44fc2-486">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-486">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-487">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-487">Requirements</span></span>

|<span data-ttu-id="44fc2-488">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-488">Requirement</span></span>|<span data-ttu-id="44fc2-489">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-490">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-490">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-491">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-491">1.0</span></span>|
|[<span data-ttu-id="44fc2-492">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-492">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-493">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-494">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-494">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-495">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-495">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-496">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-496">Example</span></span>

<span data-ttu-id="44fc2-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="44fc2-499">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="44fc2-499">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="44fc2-500">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="44fc2-500">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="44fc2-501">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-501">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-502">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-502">Type</span></span>

*   [<span data-ttu-id="44fc2-503">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="44fc2-503">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="44fc2-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-504">Requirements</span></span>

|<span data-ttu-id="44fc2-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-505">Requirement</span></span>|<span data-ttu-id="44fc2-506">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-507">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-508">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-508">1.0</span></span>|
|[<span data-ttu-id="44fc2-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-510">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-511">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-512">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-512">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-513">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-513">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="44fc2-514">local: cadeia de caracteres | [Local](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="44fc2-514">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="44fc2-515">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-515">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44fc2-516">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="44fc2-516">Read mode</span></span>

<span data-ttu-id="44fc2-517">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-517">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="44fc2-518">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="44fc2-518">Compose mode</span></span>

<span data-ttu-id="44fc2-519">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-519">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="44fc2-520">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-520">Type</span></span>

*   <span data-ttu-id="44fc2-521">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="44fc2-521">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-522">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-522">Requirements</span></span>

|<span data-ttu-id="44fc2-523">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-523">Requirement</span></span>|<span data-ttu-id="44fc2-524">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-525">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-526">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-526">1.0</span></span>|
|[<span data-ttu-id="44fc2-527">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-527">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-528">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-529">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-530">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-530">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="44fc2-531">normalizedSubject: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="44fc2-531">normalizedSubject: String</span></span>

<span data-ttu-id="44fc2-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="44fc2-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="44fc2-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-536">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-536">Type</span></span>

*   <span data-ttu-id="44fc2-537">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-537">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-538">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-538">Requirements</span></span>

|<span data-ttu-id="44fc2-539">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-539">Requirement</span></span>|<span data-ttu-id="44fc2-540">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-541">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-542">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-542">1.0</span></span>|
|[<span data-ttu-id="44fc2-543">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-544">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-545">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-546">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-547">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-547">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="44fc2-548">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="44fc2-548">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="44fc2-549">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-549">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-550">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-550">Type</span></span>

*   [<span data-ttu-id="44fc2-551">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="44fc2-551">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="44fc2-552">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-552">Requirements</span></span>

|<span data-ttu-id="44fc2-553">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-553">Requirement</span></span>|<span data-ttu-id="44fc2-554">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-554">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-555">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-555">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-556">1.3</span><span class="sxs-lookup"><span data-stu-id="44fc2-556">1.3</span></span>|
|[<span data-ttu-id="44fc2-557">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-557">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-558">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-558">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-559">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-559">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-560">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-560">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-561">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-561">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="44fc2-562">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[](/javascript/api/outlook/office.recipients) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="44fc2-562">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="44fc2-563">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="44fc2-563">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="44fc2-564">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="44fc2-564">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44fc2-565">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="44fc2-565">Read mode</span></span>

<span data-ttu-id="44fc2-566">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="44fc2-566">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="44fc2-567">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="44fc2-567">Compose mode</span></span>

<span data-ttu-id="44fc2-568">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="44fc2-568">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="44fc2-569">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-569">Type</span></span>

*   <span data-ttu-id="44fc2-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44fc2-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-571">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-571">Requirements</span></span>

|<span data-ttu-id="44fc2-572">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-572">Requirement</span></span>|<span data-ttu-id="44fc2-573">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-573">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-574">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-575">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-575">1.0</span></span>|
|[<span data-ttu-id="44fc2-576">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-577">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-578">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-579">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-579">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="44fc2-580">organizador: [](/javascript/api/outlook/office.emailaddressdetails)|[organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="44fc2-580">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="44fc2-581">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-581">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44fc2-582">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="44fc2-582">Read mode</span></span>

<span data-ttu-id="44fc2-583">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="44fc2-583">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="44fc2-584">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="44fc2-584">Compose mode</span></span>

<span data-ttu-id="44fc2-585">A `organizer` propriedade retorna um [](/javascript/api/outlook/office.organizer) objeto organizador que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="44fc2-585">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="44fc2-586">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-586">Type</span></span>

*   <span data-ttu-id="44fc2-587">[](/javascript/api/outlook/office.emailaddressdetails) | [Organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="44fc2-587">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-588">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-588">Requirements</span></span>

|<span data-ttu-id="44fc2-589">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-589">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="44fc2-590">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-591">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-591">1.0</span></span>|<span data-ttu-id="44fc2-592">1.7</span><span class="sxs-lookup"><span data-stu-id="44fc2-592">1.7</span></span>|
|[<span data-ttu-id="44fc2-593">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-594">ReadItem</span></span>|<span data-ttu-id="44fc2-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="44fc2-596">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-597">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-597">Read</span></span>|<span data-ttu-id="44fc2-598">Escrever</span><span class="sxs-lookup"><span data-stu-id="44fc2-598">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="44fc2-599">(anulável) recorrência [](/javascript/api/outlook/office.recurrence) : recorrência</span><span class="sxs-lookup"><span data-stu-id="44fc2-599">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="44fc2-600">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-600">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="44fc2-601">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="44fc2-601">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="44fc2-602">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-602">Read and compose modes for appointment items.</span></span> <span data-ttu-id="44fc2-603">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="44fc2-603">Read mode for meeting request items.</span></span>

<span data-ttu-id="44fc2-604">A `recurrence` propriedade retorna um [](/javascript/api/outlook/office.recurrence) objeto de recorrência para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="44fc2-604">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="44fc2-605">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="44fc2-605">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="44fc2-606">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="44fc2-606">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="44fc2-607">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="44fc2-607">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="44fc2-608">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="44fc2-608">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44fc2-609">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="44fc2-609">Read mode</span></span>

<span data-ttu-id="44fc2-610">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-610">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="44fc2-611">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="44fc2-611">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="44fc2-612">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="44fc2-612">Compose mode</span></span>

<span data-ttu-id="44fc2-613">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-613">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="44fc2-614">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="44fc2-614">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="44fc2-615">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-615">Type</span></span>

* [<span data-ttu-id="44fc2-616">Recorrência</span><span class="sxs-lookup"><span data-stu-id="44fc2-616">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="44fc2-617">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-617">Requirement</span></span>|<span data-ttu-id="44fc2-618">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-619">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-620">1.7</span><span class="sxs-lookup"><span data-stu-id="44fc2-620">1.7</span></span>|
|[<span data-ttu-id="44fc2-621">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-622">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-623">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-624">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-624">Compose or Read</span></span>|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="44fc2-625">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[](/javascript/api/outlook/office.recipients) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="44fc2-625">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="44fc2-626">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="44fc2-626">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="44fc2-627">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="44fc2-627">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44fc2-628">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="44fc2-628">Read mode</span></span>

<span data-ttu-id="44fc2-629">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="44fc2-629">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="44fc2-630">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="44fc2-630">Compose mode</span></span>

<span data-ttu-id="44fc2-631">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="44fc2-631">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="44fc2-632">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-632">Type</span></span>

*   <span data-ttu-id="44fc2-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44fc2-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-634">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-634">Requirements</span></span>

|<span data-ttu-id="44fc2-635">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-635">Requirement</span></span>|<span data-ttu-id="44fc2-636">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-637">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-638">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-638">1.0</span></span>|
|[<span data-ttu-id="44fc2-639">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-640">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-641">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-642">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-642">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="44fc2-643">remetente: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="44fc2-643">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="44fc2-p128">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="44fc2-p129">As propriedades [`from`](#from-emailaddressdetailsfrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-648">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-648">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-649">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-649">Type</span></span>

*   [<span data-ttu-id="44fc2-650">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="44fc2-650">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="44fc2-651">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-651">Requirements</span></span>

|<span data-ttu-id="44fc2-652">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-652">Requirement</span></span>|<span data-ttu-id="44fc2-653">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-654">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-655">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-655">1.0</span></span>|
|[<span data-ttu-id="44fc2-656">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-657">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-658">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-659">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-659">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-660">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-660">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="44fc2-661">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="44fc2-661">(nullable) seriesId: String</span></span>

<span data-ttu-id="44fc2-662">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="44fc2-662">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="44fc2-663">No OWA e no Outlook, `seriesId` o retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="44fc2-663">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="44fc2-664">No entanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="44fc2-664">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-665">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="44fc2-665">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="44fc2-666">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="44fc2-666">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="44fc2-667">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="44fc2-667">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="44fc2-668">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="44fc2-668">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="44fc2-669">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="44fc2-669">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="44fc2-670">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-670">Type</span></span>

* <span data-ttu-id="44fc2-671">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-671">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-672">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-672">Requirements</span></span>

|<span data-ttu-id="44fc2-673">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-673">Requirement</span></span>|<span data-ttu-id="44fc2-674">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-674">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-675">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-675">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-676">1.7</span><span class="sxs-lookup"><span data-stu-id="44fc2-676">1.7</span></span>|
|[<span data-ttu-id="44fc2-677">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-677">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-678">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-678">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-679">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-679">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-680">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-680">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-681">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-681">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="44fc2-682">Início: data | [Tempo](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="44fc2-682">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="44fc2-683">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="44fc2-683">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="44fc2-p132">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44fc2-686">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="44fc2-686">Read mode</span></span>

<span data-ttu-id="44fc2-687">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-687">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="44fc2-688">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="44fc2-688">Compose mode</span></span>

<span data-ttu-id="44fc2-689">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-689">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="44fc2-690">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="44fc2-690">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="44fc2-691">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-691">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="44fc2-692">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-692">Type</span></span>

*   <span data-ttu-id="44fc2-693">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="44fc2-693">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-694">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-694">Requirements</span></span>

|<span data-ttu-id="44fc2-695">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-695">Requirement</span></span>|<span data-ttu-id="44fc2-696">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-697">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-698">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-698">1.0</span></span>|
|[<span data-ttu-id="44fc2-699">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-700">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-701">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-702">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-702">Compose or Read</span></span>|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="44fc2-703">subject: cadeia de caracteres | [Assunto](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="44fc2-703">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="44fc2-704">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-704">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="44fc2-705">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="44fc2-705">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44fc2-706">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="44fc2-706">Read mode</span></span>

<span data-ttu-id="44fc2-p133">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="44fc2-709">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="44fc2-709">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="44fc2-710">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="44fc2-710">Compose mode</span></span>
<span data-ttu-id="44fc2-711">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="44fc2-711">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="44fc2-712">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-712">Type</span></span>

*   <span data-ttu-id="44fc2-713">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="44fc2-713">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-714">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-714">Requirements</span></span>

|<span data-ttu-id="44fc2-715">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-715">Requirement</span></span>|<span data-ttu-id="44fc2-716">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-716">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-717">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-717">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-718">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-718">1.0</span></span>|
|[<span data-ttu-id="44fc2-719">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-719">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-720">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-720">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-721">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-721">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-722">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-722">Compose or Read</span></span>|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="44fc2-723">para: Array. <[](/javascript/api/outlook/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook/office.recipients) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="44fc2-723">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="44fc2-724">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="44fc2-724">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="44fc2-725">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="44fc2-725">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44fc2-726">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="44fc2-726">Read mode</span></span>

<span data-ttu-id="44fc2-p135">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="44fc2-729">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="44fc2-729">Compose mode</span></span>

<span data-ttu-id="44fc2-730">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="44fc2-730">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="44fc2-731">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-731">Type</span></span>

*   <span data-ttu-id="44fc2-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44fc2-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-733">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-733">Requirements</span></span>

|<span data-ttu-id="44fc2-734">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-734">Requirement</span></span>|<span data-ttu-id="44fc2-735">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-735">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-736">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-736">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-737">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-737">1.0</span></span>|
|[<span data-ttu-id="44fc2-738">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-738">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-739">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-739">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-740">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-740">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-741">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-741">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="44fc2-742">Métodos</span><span class="sxs-lookup"><span data-stu-id="44fc2-742">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="44fc2-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="44fc2-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="44fc2-744">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="44fc2-744">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="44fc2-745">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="44fc2-745">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="44fc2-746">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="44fc2-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-747">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-747">Parameters</span></span>
|<span data-ttu-id="44fc2-748">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-748">Name</span></span>|<span data-ttu-id="44fc2-749">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-749">Type</span></span>|<span data-ttu-id="44fc2-750">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-750">Attributes</span></span>|<span data-ttu-id="44fc2-751">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-751">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="44fc2-752">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-752">String</span></span>||<span data-ttu-id="44fc2-p136">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="44fc2-755">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-755">String</span></span>||<span data-ttu-id="44fc2-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="44fc2-758">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-758">Object</span></span>|<span data-ttu-id="44fc2-759">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-759">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-760">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44fc2-761">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-761">Object</span></span>|<span data-ttu-id="44fc2-762">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-762">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-763">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="44fc2-764">Booliano</span><span class="sxs-lookup"><span data-stu-id="44fc2-764">Boolean</span></span>|<span data-ttu-id="44fc2-765">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-765">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-766">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="44fc2-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="44fc2-767">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-767">function</span></span>|<span data-ttu-id="44fc2-768">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-768">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-769">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="44fc2-770">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="44fc2-771">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="44fc2-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="44fc2-772">Erros</span><span class="sxs-lookup"><span data-stu-id="44fc2-772">Errors</span></span>

|<span data-ttu-id="44fc2-773">Código de erro</span><span class="sxs-lookup"><span data-stu-id="44fc2-773">Error code</span></span>|<span data-ttu-id="44fc2-774">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="44fc2-775">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="44fc2-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="44fc2-776">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="44fc2-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="44fc2-777">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="44fc2-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-778">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-778">Requirements</span></span>

|<span data-ttu-id="44fc2-779">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-779">Requirement</span></span>|<span data-ttu-id="44fc2-780">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-781">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-782">1.1</span><span class="sxs-lookup"><span data-stu-id="44fc2-782">1.1</span></span>|
|[<span data-ttu-id="44fc2-783">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="44fc2-785">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-786">Escrever</span><span class="sxs-lookup"><span data-stu-id="44fc2-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="44fc2-787">Exemplos</span><span class="sxs-lookup"><span data-stu-id="44fc2-787">Examples</span></span>

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

<span data-ttu-id="44fc2-788">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="44fc2-788">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="44fc2-789">addFileAttachmentFromBase64Async (base64file, AttachmentName, [Options], [callback])</span><span class="sxs-lookup"><span data-stu-id="44fc2-789">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="44fc2-790">Adiciona um arquivo da codificação Base64 a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="44fc2-790">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="44fc2-791">O `addFileAttachmentFromBase64Async` método carrega o arquivo da codificação Base64 e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="44fc2-791">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="44fc2-792">Esse método retorna o identificador de anexo no objeto AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="44fc2-792">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="44fc2-793">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="44fc2-793">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-794">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-794">Parameters</span></span>

|<span data-ttu-id="44fc2-795">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-795">Name</span></span>|<span data-ttu-id="44fc2-796">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-796">Type</span></span>|<span data-ttu-id="44fc2-797">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-797">Attributes</span></span>|<span data-ttu-id="44fc2-798">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-798">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="44fc2-799">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-799">String</span></span>||<span data-ttu-id="44fc2-800">O conteúdo codificado em Base64 de uma imagem ou arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="44fc2-800">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="44fc2-801">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-801">String</span></span>||<span data-ttu-id="44fc2-p139">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="44fc2-804">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-804">Object</span></span>|<span data-ttu-id="44fc2-805">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-805">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-806">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-806">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44fc2-807">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-807">Object</span></span>|<span data-ttu-id="44fc2-808">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-808">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-809">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-809">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="44fc2-810">Booliano</span><span class="sxs-lookup"><span data-stu-id="44fc2-810">Boolean</span></span>|<span data-ttu-id="44fc2-811">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-811">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-812">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="44fc2-812">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="44fc2-813">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-813">function</span></span>|<span data-ttu-id="44fc2-814">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-814">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-815">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-815">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="44fc2-816">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-816">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="44fc2-817">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="44fc2-817">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="44fc2-818">Erros</span><span class="sxs-lookup"><span data-stu-id="44fc2-818">Errors</span></span>

|<span data-ttu-id="44fc2-819">Código de erro</span><span class="sxs-lookup"><span data-stu-id="44fc2-819">Error code</span></span>|<span data-ttu-id="44fc2-820">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-820">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="44fc2-821">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="44fc2-821">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="44fc2-822">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="44fc2-822">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="44fc2-823">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="44fc2-823">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-824">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-824">Requirements</span></span>

|<span data-ttu-id="44fc2-825">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-825">Requirement</span></span>|<span data-ttu-id="44fc2-826">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-827">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-828">Visualização</span><span class="sxs-lookup"><span data-stu-id="44fc2-828">Preview</span></span>|
|[<span data-ttu-id="44fc2-829">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-830">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-830">ReadWriteItem</span></span>|
|[<span data-ttu-id="44fc2-831">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-832">Escrever</span><span class="sxs-lookup"><span data-stu-id="44fc2-832">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="44fc2-833">Exemplos</span><span class="sxs-lookup"><span data-stu-id="44fc2-833">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="44fc2-834">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="44fc2-834">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="44fc2-835">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="44fc2-835">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="44fc2-836">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="44fc2-836">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-837">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-837">Parameters</span></span>

| <span data-ttu-id="44fc2-838">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-838">Name</span></span> | <span data-ttu-id="44fc2-839">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-839">Type</span></span> | <span data-ttu-id="44fc2-840">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-840">Attributes</span></span> | <span data-ttu-id="44fc2-841">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-841">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="44fc2-842">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="44fc2-842">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="44fc2-843">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="44fc2-843">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="44fc2-844">Função</span><span class="sxs-lookup"><span data-stu-id="44fc2-844">Function</span></span> || <span data-ttu-id="44fc2-p140">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="44fc2-848">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-848">Object</span></span> | <span data-ttu-id="44fc2-849">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-849">&lt;optional&gt;</span></span> | <span data-ttu-id="44fc2-850">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-850">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="44fc2-851">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-851">Object</span></span> | <span data-ttu-id="44fc2-852">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-852">&lt;optional&gt;</span></span> | <span data-ttu-id="44fc2-853">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-853">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="44fc2-854">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-854">function</span></span>| <span data-ttu-id="44fc2-855">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-855">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-856">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-857">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-857">Requirements</span></span>

|<span data-ttu-id="44fc2-858">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-858">Requirement</span></span>| <span data-ttu-id="44fc2-859">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-859">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-860">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="44fc2-861">1.7</span><span class="sxs-lookup"><span data-stu-id="44fc2-861">1.7</span></span> |
|[<span data-ttu-id="44fc2-862">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="44fc2-863">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-863">ReadItem</span></span> |
|[<span data-ttu-id="44fc2-864">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="44fc2-865">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-865">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="44fc2-866">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-866">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="44fc2-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="44fc2-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="44fc2-868">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-868">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="44fc2-p141">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="44fc2-872">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="44fc2-872">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="44fc2-873">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="44fc2-873">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-874">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-874">Parameters</span></span>

|<span data-ttu-id="44fc2-875">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-875">Name</span></span>|<span data-ttu-id="44fc2-876">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-876">Type</span></span>|<span data-ttu-id="44fc2-877">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-877">Attributes</span></span>|<span data-ttu-id="44fc2-878">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-878">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="44fc2-879">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-879">String</span></span>||<span data-ttu-id="44fc2-p142">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="44fc2-882">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="44fc2-882">String</span></span>||<span data-ttu-id="44fc2-883">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="44fc2-883">The subject of the item to be attached.</span></span> <span data-ttu-id="44fc2-884">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="44fc2-884">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="44fc2-885">Object</span><span class="sxs-lookup"><span data-stu-id="44fc2-885">Object</span></span>|<span data-ttu-id="44fc2-886">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-886">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-887">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-887">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44fc2-888">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-888">Object</span></span>|<span data-ttu-id="44fc2-889">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-889">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-890">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-890">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44fc2-891">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-891">function</span></span>|<span data-ttu-id="44fc2-892">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-892">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-893">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-893">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="44fc2-894">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-894">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="44fc2-895">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="44fc2-895">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="44fc2-896">Erros</span><span class="sxs-lookup"><span data-stu-id="44fc2-896">Errors</span></span>

|<span data-ttu-id="44fc2-897">Código de erro</span><span class="sxs-lookup"><span data-stu-id="44fc2-897">Error code</span></span>|<span data-ttu-id="44fc2-898">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-898">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="44fc2-899">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="44fc2-899">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-900">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-900">Requirements</span></span>

|<span data-ttu-id="44fc2-901">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-901">Requirement</span></span>|<span data-ttu-id="44fc2-902">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-903">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-904">1.1</span><span class="sxs-lookup"><span data-stu-id="44fc2-904">1.1</span></span>|
|[<span data-ttu-id="44fc2-905">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="44fc2-907">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-908">Escrever</span><span class="sxs-lookup"><span data-stu-id="44fc2-908">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-909">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-909">Example</span></span>

<span data-ttu-id="44fc2-910">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-910">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="44fc2-911">close()</span><span class="sxs-lookup"><span data-stu-id="44fc2-911">close()</span></span>

<span data-ttu-id="44fc2-912">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="44fc2-912">Closes the current item that is being composed.</span></span>

<span data-ttu-id="44fc2-p144">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-915">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="44fc2-915">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="44fc2-916">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="44fc2-916">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-917">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-917">Requirements</span></span>

|<span data-ttu-id="44fc2-918">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-918">Requirement</span></span>|<span data-ttu-id="44fc2-919">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-919">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-920">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-920">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-921">1.3</span><span class="sxs-lookup"><span data-stu-id="44fc2-921">1.3</span></span>|
|[<span data-ttu-id="44fc2-922">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-922">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-923">Restrito</span><span class="sxs-lookup"><span data-stu-id="44fc2-923">Restricted</span></span>|
|[<span data-ttu-id="44fc2-924">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-924">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-925">Escrever</span><span class="sxs-lookup"><span data-stu-id="44fc2-925">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="44fc2-926">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="44fc2-926">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="44fc2-927">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="44fc2-927">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-928">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="44fc2-928">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="44fc2-929">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="44fc2-929">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="44fc2-930">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="44fc2-930">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="44fc2-p145">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-934">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-934">Parameters</span></span>

|<span data-ttu-id="44fc2-935">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-935">Name</span></span>|<span data-ttu-id="44fc2-936">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-936">Type</span></span>|<span data-ttu-id="44fc2-937">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-937">Attributes</span></span>|<span data-ttu-id="44fc2-938">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-938">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="44fc2-939">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="44fc2-939">String &#124; Object</span></span>||<span data-ttu-id="44fc2-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="44fc2-942">**OU**</span><span class="sxs-lookup"><span data-stu-id="44fc2-942">**OR**</span></span><br/><span data-ttu-id="44fc2-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="44fc2-945">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-945">String</span></span>|<span data-ttu-id="44fc2-946">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-946">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="44fc2-949">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-949">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="44fc2-950">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-950">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-951">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-951">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="44fc2-952">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-952">String</span></span>||<span data-ttu-id="44fc2-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="44fc2-955">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="44fc2-955">String</span></span>||<span data-ttu-id="44fc2-956">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="44fc2-956">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="44fc2-957">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-957">String</span></span>||<span data-ttu-id="44fc2-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="44fc2-960">Booliano</span><span class="sxs-lookup"><span data-stu-id="44fc2-960">Boolean</span></span>||<span data-ttu-id="44fc2-p151">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="44fc2-963">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-963">String</span></span>||<span data-ttu-id="44fc2-p152">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="44fc2-967">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-967">function</span></span>|<span data-ttu-id="44fc2-968">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-968">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-969">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-969">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-970">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-970">Requirements</span></span>

|<span data-ttu-id="44fc2-971">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-971">Requirement</span></span>|<span data-ttu-id="44fc2-972">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-972">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-973">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-973">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-974">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-974">1.0</span></span>|
|[<span data-ttu-id="44fc2-975">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-975">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-976">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-976">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-977">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-977">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-978">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-978">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="44fc2-979">Exemplos</span><span class="sxs-lookup"><span data-stu-id="44fc2-979">Examples</span></span>

<span data-ttu-id="44fc2-980">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-980">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="44fc2-981">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="44fc2-981">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="44fc2-982">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="44fc2-982">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="44fc2-983">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="44fc2-983">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="44fc2-984">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-984">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="44fc2-985">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-985">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="44fc2-986">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="44fc2-986">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="44fc2-987">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="44fc2-987">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-988">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="44fc2-988">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="44fc2-989">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="44fc2-989">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="44fc2-990">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="44fc2-990">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="44fc2-p153">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-994">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-994">Parameters</span></span>

|<span data-ttu-id="44fc2-995">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-995">Name</span></span>|<span data-ttu-id="44fc2-996">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-996">Type</span></span>|<span data-ttu-id="44fc2-997">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-997">Attributes</span></span>|<span data-ttu-id="44fc2-998">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-998">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="44fc2-999">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="44fc2-999">String &#124; Object</span></span>||<span data-ttu-id="44fc2-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="44fc2-1002">**OU**</span><span class="sxs-lookup"><span data-stu-id="44fc2-1002">**OR**</span></span><br/><span data-ttu-id="44fc2-p155">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="44fc2-1005">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-1005">String</span></span>|<span data-ttu-id="44fc2-1006">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-p156">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="44fc2-1009">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1009">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="44fc2-1010">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1011">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1011">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="44fc2-1012">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-1012">String</span></span>||<span data-ttu-id="44fc2-p157">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="44fc2-1015">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="44fc2-1015">String</span></span>||<span data-ttu-id="44fc2-1016">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1016">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="44fc2-1017">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-1017">String</span></span>||<span data-ttu-id="44fc2-p158">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="44fc2-1020">Booliano</span><span class="sxs-lookup"><span data-stu-id="44fc2-1020">Boolean</span></span>||<span data-ttu-id="44fc2-p159">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="44fc2-1023">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-1023">String</span></span>||<span data-ttu-id="44fc2-p160">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="44fc2-1027">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-1027">function</span></span>|<span data-ttu-id="44fc2-1028">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1028">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1029">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1029">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1030">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1030">Requirements</span></span>

|<span data-ttu-id="44fc2-1031">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1031">Requirement</span></span>|<span data-ttu-id="44fc2-1032">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1032">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1033">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1033">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1034">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-1034">1.0</span></span>|
|[<span data-ttu-id="44fc2-1035">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1035">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1036">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1036">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-1037">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-1037">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1038">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-1038">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="44fc2-1039">Exemplos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1039">Examples</span></span>

<span data-ttu-id="44fc2-1040">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1040">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="44fc2-1041">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1041">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="44fc2-1042">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1042">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="44fc2-1043">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1043">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="44fc2-1044">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1044">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="44fc2-1045">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1045">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="44fc2-1046">getAttachmentContentAsync (attachmentid, [opções], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="44fc2-1046">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="44fc2-1047">Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um `AttachmentContent` objeto.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1047">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="44fc2-1048">O `getAttachmentContentAsync` método obtém o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1048">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="44fc2-1049">Como prática recomendada, você deve usar o identificador para recuperar um anexo na mesma sessão em que o attachmentIds foi recuperado com a `getAttachmentsAsync` chamada ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="44fc2-1049">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="44fc2-1050">No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1050">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="44fc2-1051">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1051">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1052">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1052">Parameters</span></span>

|<span data-ttu-id="44fc2-1053">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1053">Name</span></span>|<span data-ttu-id="44fc2-1054">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1054">Type</span></span>|<span data-ttu-id="44fc2-1055">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1055">Attributes</span></span>|<span data-ttu-id="44fc2-1056">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1056">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="44fc2-1057">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-1057">String</span></span>||<span data-ttu-id="44fc2-1058">O identificador do anexo que você deseja obter.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1058">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="44fc2-1059">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1059">Object</span></span>|<span data-ttu-id="44fc2-1060">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1060">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1061">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1061">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44fc2-1062">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1062">Object</span></span>|<span data-ttu-id="44fc2-1063">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1064">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1064">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44fc2-1065">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-1065">function</span></span>|<span data-ttu-id="44fc2-1066">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1067">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1067">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1068">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1068">Requirements</span></span>

|<span data-ttu-id="44fc2-1069">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1069">Requirement</span></span>|<span data-ttu-id="44fc2-1070">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1070">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1071">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1071">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1072">Visualização</span><span class="sxs-lookup"><span data-stu-id="44fc2-1072">Preview</span></span>|
|[<span data-ttu-id="44fc2-1073">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1073">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1074">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1074">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-1075">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-1075">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1076">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-1076">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44fc2-1077">Retorna:</span><span class="sxs-lookup"><span data-stu-id="44fc2-1077">Returns:</span></span>

<span data-ttu-id="44fc2-1078">Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="44fc2-1078">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="44fc2-1079">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1079">Example</span></span>

```javascript
var item = Office.context.mailbox.item;
var listOfAttachments = [];
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var options = {asyncContext: {type: result.value[i].attachmentType}};
      getAttachmentContentAsync(result.value[i].id, options, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  if (result.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    // Handle file attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
    // Handle email item attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
    // Handle .icalender attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    // Handle cloud attachment.
  } else {
    // Handle attachment formats that are not supported.
  }
}
```

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="44fc2-1080">getAttachmentsAsync ([Options], [callback]) → array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="44fc2-1080">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="44fc2-1081">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1081">Gets the item's attachments as an array.</span></span> <span data-ttu-id="44fc2-1082">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1082">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1083">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1083">Parameters</span></span>

|<span data-ttu-id="44fc2-1084">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1084">Name</span></span>|<span data-ttu-id="44fc2-1085">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1085">Type</span></span>|<span data-ttu-id="44fc2-1086">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1086">Attributes</span></span>|<span data-ttu-id="44fc2-1087">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1087">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="44fc2-1088">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1088">Object</span></span>|<span data-ttu-id="44fc2-1089">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1089">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1090">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1090">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44fc2-1091">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1091">Object</span></span>|<span data-ttu-id="44fc2-1092">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1093">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1093">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44fc2-1094">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-1094">function</span></span>|<span data-ttu-id="44fc2-1095">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1096">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1096">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1097">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1097">Requirements</span></span>

|<span data-ttu-id="44fc2-1098">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1098">Requirement</span></span>|<span data-ttu-id="44fc2-1099">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1099">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1100">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1100">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1101">Visualização</span><span class="sxs-lookup"><span data-stu-id="44fc2-1101">Preview</span></span>|
|[<span data-ttu-id="44fc2-1102">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1102">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1103">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1103">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-1104">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-1104">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1105">Escrever</span><span class="sxs-lookup"><span data-stu-id="44fc2-1105">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="44fc2-1106">Retorna:</span><span class="sxs-lookup"><span data-stu-id="44fc2-1106">Returns:</span></span>

<span data-ttu-id="44fc2-1107">Tipo: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="44fc2-1107">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="44fc2-1108">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1108">Example</span></span>

<span data-ttu-id="44fc2-1109">O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1109">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="44fc2-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="44fc2-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="44fc2-1111">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1111">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-1112">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1112">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-1113">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1113">Requirements</span></span>

|<span data-ttu-id="44fc2-1114">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1114">Requirement</span></span>|<span data-ttu-id="44fc2-1115">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1115">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1116">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1116">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1117">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-1117">1.0</span></span>|
|[<span data-ttu-id="44fc2-1118">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1118">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1119">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1119">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-1120">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-1120">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1121">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-1121">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44fc2-1122">Retorna:</span><span class="sxs-lookup"><span data-stu-id="44fc2-1122">Returns:</span></span>

<span data-ttu-id="44fc2-1123">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="44fc2-1123">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="44fc2-1124">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1124">Example</span></span>

<span data-ttu-id="44fc2-1125">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1125">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="44fc2-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="44fc2-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="44fc2-1127">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1127">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-1128">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1128">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1129">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1129">Parameters</span></span>

|<span data-ttu-id="44fc2-1130">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1130">Name</span></span>|<span data-ttu-id="44fc2-1131">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1131">Type</span></span>|<span data-ttu-id="44fc2-1132">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1132">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="44fc2-1133">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="44fc2-1133">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="44fc2-1134">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1134">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1135">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1135">Requirements</span></span>

|<span data-ttu-id="44fc2-1136">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1136">Requirement</span></span>|<span data-ttu-id="44fc2-1137">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1138">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1139">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-1139">1.0</span></span>|
|[<span data-ttu-id="44fc2-1140">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1141">Restrito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1141">Restricted</span></span>|
|[<span data-ttu-id="44fc2-1142">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1143">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44fc2-1144">Retorna:</span><span class="sxs-lookup"><span data-stu-id="44fc2-1144">Returns:</span></span>

<span data-ttu-id="44fc2-1145">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1145">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="44fc2-1146">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1146">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="44fc2-1147">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1147">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="44fc2-1148">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1148">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="44fc2-1149">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="44fc2-1149">Value of `entityType`</span></span>|<span data-ttu-id="44fc2-1150">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="44fc2-1150">Type of objects in returned array</span></span>|<span data-ttu-id="44fc2-1151">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="44fc2-1151">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="44fc2-1152">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-1152">String</span></span>|<span data-ttu-id="44fc2-1153">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="44fc2-1153">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="44fc2-1154">Contato</span><span class="sxs-lookup"><span data-stu-id="44fc2-1154">Contact</span></span>|<span data-ttu-id="44fc2-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="44fc2-1155">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="44fc2-1156">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-1156">String</span></span>|<span data-ttu-id="44fc2-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="44fc2-1157">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="44fc2-1158">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="44fc2-1158">MeetingSuggestion</span></span>|<span data-ttu-id="44fc2-1159">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="44fc2-1159">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="44fc2-1160">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="44fc2-1160">PhoneNumber</span></span>|<span data-ttu-id="44fc2-1161">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="44fc2-1161">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="44fc2-1162">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="44fc2-1162">TaskSuggestion</span></span>|<span data-ttu-id="44fc2-1163">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="44fc2-1163">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="44fc2-1164">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-1164">String</span></span>|<span data-ttu-id="44fc2-1165">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="44fc2-1165">**Restricted**</span></span>|

<span data-ttu-id="44fc2-1166">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="44fc2-1166">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="44fc2-1167">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1167">Example</span></span>

<span data-ttu-id="44fc2-1168">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1168">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="44fc2-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="44fc2-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="44fc2-1170">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1170">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-1171">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1171">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="44fc2-1172">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1172">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1173">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1173">Parameters</span></span>

|<span data-ttu-id="44fc2-1174">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1174">Name</span></span>|<span data-ttu-id="44fc2-1175">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1175">Type</span></span>|<span data-ttu-id="44fc2-1176">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1176">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="44fc2-1177">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-1177">String</span></span>|<span data-ttu-id="44fc2-1178">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1178">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1179">Requirements</span></span>

|<span data-ttu-id="44fc2-1180">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1180">Requirement</span></span>|<span data-ttu-id="44fc2-1181">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1181">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1183">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-1183">1.0</span></span>|
|[<span data-ttu-id="44fc2-1184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1185">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-1186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-1186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1187">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-1187">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44fc2-1188">Retorna:</span><span class="sxs-lookup"><span data-stu-id="44fc2-1188">Returns:</span></span>

<span data-ttu-id="44fc2-p164">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="44fc2-1191">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="44fc2-1191">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="44fc2-1192">getInitializationContextAsync ([opções], [callback])</span><span class="sxs-lookup"><span data-stu-id="44fc2-1192">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="44fc2-1193">Obtém dados de inicialização passados quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1193">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-1194">Este método só é compatível com o Outlook 2016 ou posterior no Windows (clique para executar versões posteriores a 16.0.8413.1000) e Outlook na Web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1194">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1195">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1195">Parameters</span></span>

|<span data-ttu-id="44fc2-1196">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1196">Name</span></span>|<span data-ttu-id="44fc2-1197">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1197">Type</span></span>|<span data-ttu-id="44fc2-1198">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1198">Attributes</span></span>|<span data-ttu-id="44fc2-1199">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1199">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="44fc2-1200">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1200">Object</span></span>|<span data-ttu-id="44fc2-1201">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1201">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1202">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1202">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44fc2-1203">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1203">Object</span></span>|<span data-ttu-id="44fc2-1204">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1204">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1205">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1205">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44fc2-1206">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-1206">function</span></span>|<span data-ttu-id="44fc2-1207">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1207">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1208">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1208">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="44fc2-1209">Com êxito, os dados de inicialização são fornecidos na `asyncResult.value` Propriedade como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1209">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="44fc2-1210">Se não houver nenhum contexto de inicialização, `asyncResult` o objeto conterá `Error` um objeto com `code` sua propriedade definida `9020` como e `name` sua propriedade definida `GenericResponseError`como.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1210">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1211">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1211">Requirements</span></span>

|<span data-ttu-id="44fc2-1212">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1212">Requirement</span></span>|<span data-ttu-id="44fc2-1213">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1213">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1214">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1215">Visualização</span><span class="sxs-lookup"><span data-stu-id="44fc2-1215">Preview</span></span>|
|[<span data-ttu-id="44fc2-1216">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1217">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-1218">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-1218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1219">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-1219">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-1220">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1220">Example</span></span>

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

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="44fc2-1221">getItemIdAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="44fc2-1221">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="44fc2-1222">Obtém de forma assíncrona a ID de um item salvo.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1222">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="44fc2-1223">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1223">Compose mode only.</span></span>

<span data-ttu-id="44fc2-1224">Quando invocado, este método retorna a ID do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1224">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-1225">Se seu suplemento chamar `getItemIdAsync` um item no modo de redação (por exemplo, para `itemId` usar com o EWS ou a API REST), lembre-se de que, quando o Outlook estiver no modo cache, pode levar algum tempo para que o item seja sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1225">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="44fc2-1226">Até que o item seja sincronizado, `itemId` o não é reconhecido e usado retorna um erro.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1226">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1227">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1227">Parameters</span></span>

|<span data-ttu-id="44fc2-1228">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1228">Name</span></span>|<span data-ttu-id="44fc2-1229">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1229">Type</span></span>|<span data-ttu-id="44fc2-1230">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1230">Attributes</span></span>|<span data-ttu-id="44fc2-1231">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1231">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="44fc2-1232">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1232">Object</span></span>|<span data-ttu-id="44fc2-1233">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1233">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1234">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1234">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44fc2-1235">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1235">Object</span></span>|<span data-ttu-id="44fc2-1236">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1236">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1237">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1237">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44fc2-1238">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-1238">function</span></span>||<span data-ttu-id="44fc2-1239">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1239">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="44fc2-1240">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1240">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="44fc2-1241">Erros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1241">Errors</span></span>

|<span data-ttu-id="44fc2-1242">Código de erro</span><span class="sxs-lookup"><span data-stu-id="44fc2-1242">Error code</span></span>|<span data-ttu-id="44fc2-1243">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1243">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="44fc2-1244">A ID não pode ser recuperada até que o item seja salvo.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1244">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1245">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1245">Requirements</span></span>

|<span data-ttu-id="44fc2-1246">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1246">Requirement</span></span>|<span data-ttu-id="44fc2-1247">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1247">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1248">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1248">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1249">Visualização</span><span class="sxs-lookup"><span data-stu-id="44fc2-1249">Preview</span></span>|
|[<span data-ttu-id="44fc2-1250">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1250">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1251">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-1252">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-1252">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1253">Escrever</span><span class="sxs-lookup"><span data-stu-id="44fc2-1253">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="44fc2-1254">Exemplos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1254">Examples</span></span>

```javascript
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="44fc2-1255">O exemplo a seguir mostra a estrutura do `result` parâmetro que é passado para a função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1255">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="44fc2-1256">A `value` propriedade contém a ID do item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1256">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="44fc2-1257">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="44fc2-1257">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="44fc2-1258">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1258">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-1259">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1259">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="44fc2-p168">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="44fc2-1263">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="44fc2-1263">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="44fc2-1264">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1264">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="44fc2-p169">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-1268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1268">Requirements</span></span>

|<span data-ttu-id="44fc2-1269">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1269">Requirement</span></span>|<span data-ttu-id="44fc2-1270">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1271">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-1272">1.0</span></span>|
|[<span data-ttu-id="44fc2-1273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1274">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-1275">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-1275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1276">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44fc2-1277">Retorna:</span><span class="sxs-lookup"><span data-stu-id="44fc2-1277">Returns:</span></span>

<span data-ttu-id="44fc2-p170">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="44fc2-1280">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="44fc2-1280">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="44fc2-1281">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1281">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="44fc2-1282">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1282">Example</span></span>

<span data-ttu-id="44fc2-1283">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1283">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="44fc2-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="44fc2-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="44fc2-1285">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1285">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-1286">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1286">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="44fc2-1287">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1287">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="44fc2-p171">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1290">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1290">Parameters</span></span>

|<span data-ttu-id="44fc2-1291">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1291">Name</span></span>|<span data-ttu-id="44fc2-1292">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1292">Type</span></span>|<span data-ttu-id="44fc2-1293">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1293">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="44fc2-1294">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-1294">String</span></span>|<span data-ttu-id="44fc2-1295">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1295">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1296">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1296">Requirements</span></span>

|<span data-ttu-id="44fc2-1297">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1297">Requirement</span></span>|<span data-ttu-id="44fc2-1298">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1298">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1299">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1299">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1300">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-1300">1.0</span></span>|
|[<span data-ttu-id="44fc2-1301">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1301">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1302">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-1303">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-1303">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1304">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-1304">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44fc2-1305">Retorna:</span><span class="sxs-lookup"><span data-stu-id="44fc2-1305">Returns:</span></span>

<span data-ttu-id="44fc2-1306">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1306">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="44fc2-1307">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="44fc2-1307">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="44fc2-1308">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="44fc2-1308">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="44fc2-1309">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1309">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="44fc2-1310">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="44fc2-1310">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="44fc2-1311">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1311">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="44fc2-p172">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p172">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1314">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1314">Parameters</span></span>

|<span data-ttu-id="44fc2-1315">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1315">Name</span></span>|<span data-ttu-id="44fc2-1316">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1316">Type</span></span>|<span data-ttu-id="44fc2-1317">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1317">Attributes</span></span>|<span data-ttu-id="44fc2-1318">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1318">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="44fc2-1319">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="44fc2-1319">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="44fc2-p173">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="44fc2-1323">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1323">Object</span></span>|<span data-ttu-id="44fc2-1324">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1324">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1325">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1325">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44fc2-1326">Object</span><span class="sxs-lookup"><span data-stu-id="44fc2-1326">Object</span></span>|<span data-ttu-id="44fc2-1327">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1327">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1328">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1328">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44fc2-1329">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-1329">function</span></span>||<span data-ttu-id="44fc2-1330">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1330">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="44fc2-1331">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1331">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="44fc2-1332">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1332">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1333">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1333">Requirements</span></span>

|<span data-ttu-id="44fc2-1334">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1334">Requirement</span></span>|<span data-ttu-id="44fc2-1335">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1335">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1336">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1337">1.2</span><span class="sxs-lookup"><span data-stu-id="44fc2-1337">1.2</span></span>|
|[<span data-ttu-id="44fc2-1338">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1338">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1339">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1339">ReadWriteItem</span></span>|
|[<span data-ttu-id="44fc2-1340">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-1340">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1341">Escrever</span><span class="sxs-lookup"><span data-stu-id="44fc2-1341">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="44fc2-1342">Retorna:</span><span class="sxs-lookup"><span data-stu-id="44fc2-1342">Returns:</span></span>

<span data-ttu-id="44fc2-1343">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1343">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="44fc2-1344">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="44fc2-1344">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="44fc2-1345">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-1345">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="44fc2-1346">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1346">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="44fc2-1347">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="44fc2-1347">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="44fc2-1348">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1348">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="44fc2-1349">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1349">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-1350">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1350">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-1351">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1351">Requirements</span></span>

|<span data-ttu-id="44fc2-1352">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1352">Requirement</span></span>|<span data-ttu-id="44fc2-1353">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1353">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1354">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1354">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1355">1.6</span><span class="sxs-lookup"><span data-stu-id="44fc2-1355">1.6</span></span>|
|[<span data-ttu-id="44fc2-1356">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1356">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1357">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-1358">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-1358">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1359">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-1359">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44fc2-1360">Retorna:</span><span class="sxs-lookup"><span data-stu-id="44fc2-1360">Returns:</span></span>

<span data-ttu-id="44fc2-1361">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="44fc2-1361">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="44fc2-1362">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1362">Example</span></span>

<span data-ttu-id="44fc2-1363">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1363">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="44fc2-1364">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="44fc2-1364">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="44fc2-p176">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="44fc2-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-1367">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1367">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="44fc2-p177">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="44fc2-1371">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="44fc2-1371">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="44fc2-1372">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1372">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="44fc2-p178">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="44fc2-1376">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1376">Requirements</span></span>

|<span data-ttu-id="44fc2-1377">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1377">Requirement</span></span>|<span data-ttu-id="44fc2-1378">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1378">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1379">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1380">1.6</span><span class="sxs-lookup"><span data-stu-id="44fc2-1380">1.6</span></span>|
|[<span data-ttu-id="44fc2-1381">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1381">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1382">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-1383">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-1383">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1384">Read</span><span class="sxs-lookup"><span data-stu-id="44fc2-1384">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44fc2-1385">Retorna:</span><span class="sxs-lookup"><span data-stu-id="44fc2-1385">Returns:</span></span>

<span data-ttu-id="44fc2-p179">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="44fc2-1388">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1388">Example</span></span>

<span data-ttu-id="44fc2-1389">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1389">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="44fc2-1390">getSharedPropertiesAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="44fc2-1390">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="44fc2-1391">Obtém as propriedades do compromisso ou da mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1391">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1392">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1392">Parameters</span></span>

|<span data-ttu-id="44fc2-1393">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1393">Name</span></span>|<span data-ttu-id="44fc2-1394">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1394">Type</span></span>|<span data-ttu-id="44fc2-1395">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1395">Attributes</span></span>|<span data-ttu-id="44fc2-1396">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1396">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="44fc2-1397">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1397">Object</span></span>|<span data-ttu-id="44fc2-1398">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1398">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1399">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1399">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44fc2-1400">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1400">Object</span></span>|<span data-ttu-id="44fc2-1401">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1401">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1402">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1402">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44fc2-1403">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-1403">function</span></span>||<span data-ttu-id="44fc2-1404">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1404">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="44fc2-1405">As propriedades compartilhadas são fornecidas [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) como um objeto `asyncResult.value` na propriedade.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1405">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="44fc2-1406">Este objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1406">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1407">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1407">Requirements</span></span>

|<span data-ttu-id="44fc2-1408">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1408">Requirement</span></span>|<span data-ttu-id="44fc2-1409">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1409">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1410">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1411">Visualização</span><span class="sxs-lookup"><span data-stu-id="44fc2-1411">Preview</span></span>|
|[<span data-ttu-id="44fc2-1412">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1413">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-1414">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-1414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1415">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-1415">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-1416">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1416">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="44fc2-1417">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="44fc2-1417">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="44fc2-1418">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1418">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="44fc2-p181">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p181">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1422">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1422">Parameters</span></span>

|<span data-ttu-id="44fc2-1423">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1423">Name</span></span>|<span data-ttu-id="44fc2-1424">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1424">Type</span></span>|<span data-ttu-id="44fc2-1425">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1425">Attributes</span></span>|<span data-ttu-id="44fc2-1426">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1426">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="44fc2-1427">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-1427">function</span></span>||<span data-ttu-id="44fc2-1428">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1428">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="44fc2-1429">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1429">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="44fc2-1430">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1430">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="44fc2-1431">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1431">Object</span></span>|<span data-ttu-id="44fc2-1432">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1432">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1433">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1433">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="44fc2-1434">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1434">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1435">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1435">Requirements</span></span>

|<span data-ttu-id="44fc2-1436">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1436">Requirement</span></span>|<span data-ttu-id="44fc2-1437">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1437">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1438">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1439">1.0</span><span class="sxs-lookup"><span data-stu-id="44fc2-1439">1.0</span></span>|
|[<span data-ttu-id="44fc2-1440">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1440">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1441">ReadItem</span></span>|
|[<span data-ttu-id="44fc2-1442">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-1442">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1443">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-1443">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-1444">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1444">Example</span></span>

<span data-ttu-id="44fc2-p184">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p184">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="44fc2-1448">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="44fc2-1448">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="44fc2-1449">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1449">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="44fc2-1450">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1450">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="44fc2-1451">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1451">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="44fc2-1452">No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1452">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="44fc2-1453">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1453">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1454">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1454">Parameters</span></span>

|<span data-ttu-id="44fc2-1455">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1455">Name</span></span>|<span data-ttu-id="44fc2-1456">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1456">Type</span></span>|<span data-ttu-id="44fc2-1457">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1457">Attributes</span></span>|<span data-ttu-id="44fc2-1458">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1458">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="44fc2-1459">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-1459">String</span></span>||<span data-ttu-id="44fc2-1460">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1460">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="44fc2-1461">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1461">Object</span></span>|<span data-ttu-id="44fc2-1462">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1462">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1463">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1463">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44fc2-1464">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1464">Object</span></span>|<span data-ttu-id="44fc2-1465">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1465">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1466">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1466">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44fc2-1467">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-1467">function</span></span>|<span data-ttu-id="44fc2-1468">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1468">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1469">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1469">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="44fc2-1470">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1470">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="44fc2-1471">Erros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1471">Errors</span></span>

|<span data-ttu-id="44fc2-1472">Código de erro</span><span class="sxs-lookup"><span data-stu-id="44fc2-1472">Error code</span></span>|<span data-ttu-id="44fc2-1473">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1473">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="44fc2-1474">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1474">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1475">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1475">Requirements</span></span>

|<span data-ttu-id="44fc2-1476">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1476">Requirement</span></span>|<span data-ttu-id="44fc2-1477">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1477">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1478">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1478">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1479">1.1</span><span class="sxs-lookup"><span data-stu-id="44fc2-1479">1.1</span></span>|
|[<span data-ttu-id="44fc2-1480">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1480">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1481">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1481">ReadWriteItem</span></span>|
|[<span data-ttu-id="44fc2-1482">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-1482">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1483">Escrever</span><span class="sxs-lookup"><span data-stu-id="44fc2-1483">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-1484">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1484">Example</span></span>

<span data-ttu-id="44fc2-1485">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1485">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="44fc2-1486">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="44fc2-1486">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="44fc2-1487">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1487">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="44fc2-1488">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1488">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1489">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1489">Parameters</span></span>

| <span data-ttu-id="44fc2-1490">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1490">Name</span></span> | <span data-ttu-id="44fc2-1491">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1491">Type</span></span> | <span data-ttu-id="44fc2-1492">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1492">Attributes</span></span> | <span data-ttu-id="44fc2-1493">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1493">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="44fc2-1494">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="44fc2-1494">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="44fc2-1495">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1495">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="44fc2-1496">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1496">Object</span></span> | <span data-ttu-id="44fc2-1497">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1497">&lt;optional&gt;</span></span> | <span data-ttu-id="44fc2-1498">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1498">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="44fc2-1499">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1499">Object</span></span> | <span data-ttu-id="44fc2-1500">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1500">&lt;optional&gt;</span></span> | <span data-ttu-id="44fc2-1501">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1501">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="44fc2-1502">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-1502">function</span></span>| <span data-ttu-id="44fc2-1503">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1503">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1504">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1505">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1505">Requirements</span></span>

|<span data-ttu-id="44fc2-1506">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1506">Requirement</span></span>| <span data-ttu-id="44fc2-1507">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1507">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1508">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="44fc2-1509">1.7</span><span class="sxs-lookup"><span data-stu-id="44fc2-1509">1.7</span></span> |
|[<span data-ttu-id="44fc2-1510">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="44fc2-1511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1511">ReadItem</span></span> |
|[<span data-ttu-id="44fc2-1512">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="44fc2-1512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="44fc2-1513">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44fc2-1513">Compose or Read</span></span> |

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="44fc2-1514">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="44fc2-1514">saveAsync([options], callback)</span></span>

<span data-ttu-id="44fc2-1515">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1515">Asynchronously saves an item.</span></span>

<span data-ttu-id="44fc2-p186">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p186">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-1519">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1519">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="44fc2-1520">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1520">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="44fc2-p188">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p188">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="44fc2-1524">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="44fc2-1524">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="44fc2-1525">O Outlook para Mac não dá suporte à gravação de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1525">Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="44fc2-1526">O `saveAsync` método falha quando chamado de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1526">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="44fc2-1527">Consulte [não é possível salvar uma reunião como rascunho no Outlook para Mac usando a API do Office js](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1527">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="44fc2-1528">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1528">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1529">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1529">Parameters</span></span>

|<span data-ttu-id="44fc2-1530">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1530">Name</span></span>|<span data-ttu-id="44fc2-1531">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1531">Type</span></span>|<span data-ttu-id="44fc2-1532">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1532">Attributes</span></span>|<span data-ttu-id="44fc2-1533">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1533">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="44fc2-1534">Object</span><span class="sxs-lookup"><span data-stu-id="44fc2-1534">Object</span></span>|<span data-ttu-id="44fc2-1535">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1535">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1536">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1536">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44fc2-1537">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1537">Object</span></span>|<span data-ttu-id="44fc2-1538">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1538">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1539">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1539">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44fc2-1540">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-1540">function</span></span>||<span data-ttu-id="44fc2-1541">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="44fc2-1542">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1542">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1543">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1543">Requirements</span></span>

|<span data-ttu-id="44fc2-1544">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1544">Requirement</span></span>|<span data-ttu-id="44fc2-1545">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1545">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1546">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1547">1.3</span><span class="sxs-lookup"><span data-stu-id="44fc2-1547">1.3</span></span>|
|[<span data-ttu-id="44fc2-1548">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1549">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1549">ReadWriteItem</span></span>|
|[<span data-ttu-id="44fc2-1550">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-1550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1551">Escrever</span><span class="sxs-lookup"><span data-stu-id="44fc2-1551">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="44fc2-1552">Exemplos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1552">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="44fc2-p190">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p190">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="44fc2-1555">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="44fc2-1555">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="44fc2-1556">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1556">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="44fc2-p191">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p191">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44fc2-1560">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="44fc2-1560">Parameters</span></span>

|<span data-ttu-id="44fc2-1561">Nome</span><span class="sxs-lookup"><span data-stu-id="44fc2-1561">Name</span></span>|<span data-ttu-id="44fc2-1562">Tipo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1562">Type</span></span>|<span data-ttu-id="44fc2-1563">Atributos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1563">Attributes</span></span>|<span data-ttu-id="44fc2-1564">Descrição</span><span class="sxs-lookup"><span data-stu-id="44fc2-1564">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="44fc2-1565">String</span><span class="sxs-lookup"><span data-stu-id="44fc2-1565">String</span></span>||<span data-ttu-id="44fc2-p192">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p192">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="44fc2-1569">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1569">Object</span></span>|<span data-ttu-id="44fc2-1570">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1570">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1571">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1571">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44fc2-1572">Objeto</span><span class="sxs-lookup"><span data-stu-id="44fc2-1572">Object</span></span>|<span data-ttu-id="44fc2-1573">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1573">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-1574">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1574">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="44fc2-1575">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="44fc2-1575">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="44fc2-1576">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="44fc2-1576">&lt;optional&gt;</span></span>|<span data-ttu-id="44fc2-p193">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p193">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="44fc2-p194">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="44fc2-p194">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="44fc2-1581">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="44fc2-1581">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="44fc2-1582">function</span><span class="sxs-lookup"><span data-stu-id="44fc2-1582">function</span></span>||<span data-ttu-id="44fc2-1583">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="44fc2-1583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44fc2-1584">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44fc2-1584">Requirements</span></span>

|<span data-ttu-id="44fc2-1585">Requisito</span><span class="sxs-lookup"><span data-stu-id="44fc2-1585">Requirement</span></span>|<span data-ttu-id="44fc2-1586">Valor</span><span class="sxs-lookup"><span data-stu-id="44fc2-1586">Value</span></span>|
|---|---|
|[<span data-ttu-id="44fc2-1587">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44fc2-1587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44fc2-1588">1.2</span><span class="sxs-lookup"><span data-stu-id="44fc2-1588">1.2</span></span>|
|[<span data-ttu-id="44fc2-1589">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1589">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44fc2-1590">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44fc2-1590">ReadWriteItem</span></span>|
|[<span data-ttu-id="44fc2-1591">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44fc2-1591">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="44fc2-1592">Escrever</span><span class="sxs-lookup"><span data-stu-id="44fc2-1592">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="44fc2-1593">Exemplo</span><span class="sxs-lookup"><span data-stu-id="44fc2-1593">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
