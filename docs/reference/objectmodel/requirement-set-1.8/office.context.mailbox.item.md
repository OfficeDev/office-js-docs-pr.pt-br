---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,8
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: bb100dd4408099789d26268743264b00d3b988ac
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629662"
---
# <a name="item"></a><span data-ttu-id="c6f5f-102">item</span><span class="sxs-lookup"><span data-stu-id="c6f5f-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="c6f5f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="c6f5f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="c6f5f-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-106">Requirements</span></span>

|<span data-ttu-id="c6f5f-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-107">Requirement</span></span>|<span data-ttu-id="c6f5f-108">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-110">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-110">1.0</span></span>|
|[<span data-ttu-id="c6f5f-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-112">Restricted</span></span>|
|[<span data-ttu-id="c6f5f-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c6f5f-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-115">Members and methods</span></span>

| <span data-ttu-id="c6f5f-116">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-116">Member</span></span> | <span data-ttu-id="c6f5f-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c6f5f-118">attachments</span><span class="sxs-lookup"><span data-stu-id="c6f5f-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="c6f5f-119">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-119">Member</span></span> |
| [<span data-ttu-id="c6f5f-120">bcc</span><span class="sxs-lookup"><span data-stu-id="c6f5f-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="c6f5f-121">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-121">Member</span></span> |
| [<span data-ttu-id="c6f5f-122">body</span><span class="sxs-lookup"><span data-stu-id="c6f5f-122">body</span></span>](#body-body) | <span data-ttu-id="c6f5f-123">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-123">Member</span></span> |
| [<span data-ttu-id="c6f5f-124">categories</span><span class="sxs-lookup"><span data-stu-id="c6f5f-124">categories</span></span>](#categories-categories) | <span data-ttu-id="c6f5f-125">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-125">Member</span></span> |
| [<span data-ttu-id="c6f5f-126">cc</span><span class="sxs-lookup"><span data-stu-id="c6f5f-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c6f5f-127">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-127">Member</span></span> |
| [<span data-ttu-id="c6f5f-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="c6f5f-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="c6f5f-129">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-129">Member</span></span> |
| [<span data-ttu-id="c6f5f-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="c6f5f-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="c6f5f-131">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-131">Member</span></span> |
| [<span data-ttu-id="c6f5f-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="c6f5f-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="c6f5f-133">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-133">Member</span></span> |
| [<span data-ttu-id="c6f5f-134">end</span><span class="sxs-lookup"><span data-stu-id="c6f5f-134">end</span></span>](#end-datetime) | <span data-ttu-id="c6f5f-135">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-135">Member</span></span> |
| [<span data-ttu-id="c6f5f-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="c6f5f-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="c6f5f-137">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-137">Member</span></span> |
| [<span data-ttu-id="c6f5f-138">from</span><span class="sxs-lookup"><span data-stu-id="c6f5f-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="c6f5f-139">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-139">Member</span></span> |
| [<span data-ttu-id="c6f5f-140">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="c6f5f-141">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-141">Member</span></span> |
| [<span data-ttu-id="c6f5f-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="c6f5f-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="c6f5f-143">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-143">Member</span></span> |
| [<span data-ttu-id="c6f5f-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="c6f5f-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="c6f5f-145">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-145">Member</span></span> |
| [<span data-ttu-id="c6f5f-146">itemId</span><span class="sxs-lookup"><span data-stu-id="c6f5f-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="c6f5f-147">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-147">Member</span></span> |
| [<span data-ttu-id="c6f5f-148">itemType</span><span class="sxs-lookup"><span data-stu-id="c6f5f-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="c6f5f-149">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-149">Member</span></span> |
| [<span data-ttu-id="c6f5f-150">location</span><span class="sxs-lookup"><span data-stu-id="c6f5f-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="c6f5f-151">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-151">Member</span></span> |
| [<span data-ttu-id="c6f5f-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="c6f5f-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="c6f5f-153">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-153">Member</span></span> |
| [<span data-ttu-id="c6f5f-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="c6f5f-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="c6f5f-155">Member</span><span class="sxs-lookup"><span data-stu-id="c6f5f-155">Member</span></span> |
| [<span data-ttu-id="c6f5f-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="c6f5f-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c6f5f-157">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-157">Member</span></span> |
| [<span data-ttu-id="c6f5f-158">organizer</span><span class="sxs-lookup"><span data-stu-id="c6f5f-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="c6f5f-159">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-159">Member</span></span> |
| [<span data-ttu-id="c6f5f-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="c6f5f-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="c6f5f-161">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-161">Member</span></span> |
| [<span data-ttu-id="c6f5f-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="c6f5f-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c6f5f-163">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-163">Member</span></span> |
| [<span data-ttu-id="c6f5f-164">sender</span><span class="sxs-lookup"><span data-stu-id="c6f5f-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="c6f5f-165">Member</span><span class="sxs-lookup"><span data-stu-id="c6f5f-165">Member</span></span> |
| [<span data-ttu-id="c6f5f-166">seriesid</span><span class="sxs-lookup"><span data-stu-id="c6f5f-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="c6f5f-167">Member</span><span class="sxs-lookup"><span data-stu-id="c6f5f-167">Member</span></span> |
| [<span data-ttu-id="c6f5f-168">start</span><span class="sxs-lookup"><span data-stu-id="c6f5f-168">start</span></span>](#start-datetime) | <span data-ttu-id="c6f5f-169">Member</span><span class="sxs-lookup"><span data-stu-id="c6f5f-169">Member</span></span> |
| [<span data-ttu-id="c6f5f-170">subject</span><span class="sxs-lookup"><span data-stu-id="c6f5f-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="c6f5f-171">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-171">Member</span></span> |
| [<span data-ttu-id="c6f5f-172">to</span><span class="sxs-lookup"><span data-stu-id="c6f5f-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c6f5f-173">Membro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-173">Member</span></span> |
| [<span data-ttu-id="c6f5f-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="c6f5f-175">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-175">Method</span></span> |
| [<span data-ttu-id="c6f5f-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="c6f5f-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="c6f5f-177">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-177">Method</span></span> |
| [<span data-ttu-id="c6f5f-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c6f5f-179">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-179">Method</span></span> |
| [<span data-ttu-id="c6f5f-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="c6f5f-181">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-181">Method</span></span> |
| [<span data-ttu-id="c6f5f-182">close</span><span class="sxs-lookup"><span data-stu-id="c6f5f-182">close</span></span>](#close) | <span data-ttu-id="c6f5f-183">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-183">Method</span></span> |
| [<span data-ttu-id="c6f5f-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="c6f5f-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="c6f5f-185">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-185">Method</span></span> |
| [<span data-ttu-id="c6f5f-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="c6f5f-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="c6f5f-187">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-187">Method</span></span> |
| [<span data-ttu-id="c6f5f-188">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-188">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="c6f5f-189">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-189">Method</span></span> |
| [<span data-ttu-id="c6f5f-190">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-190">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="c6f5f-191">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-191">Method</span></span> |
| [<span data-ttu-id="c6f5f-192">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-192">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="c6f5f-193">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-193">Method</span></span> |
| [<span data-ttu-id="c6f5f-194">getEntities</span><span class="sxs-lookup"><span data-stu-id="c6f5f-194">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="c6f5f-195">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-195">Method</span></span> |
| [<span data-ttu-id="c6f5f-196">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="c6f5f-196">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="c6f5f-197">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-197">Method</span></span> |
| [<span data-ttu-id="c6f5f-198">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="c6f5f-198">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="c6f5f-199">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-199">Method</span></span> |
| [<span data-ttu-id="c6f5f-200">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="c6f5f-201">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-201">Method</span></span> |
| [<span data-ttu-id="c6f5f-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c6f5f-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="c6f5f-203">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-203">Method</span></span> |
| [<span data-ttu-id="c6f5f-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="c6f5f-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="c6f5f-205">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-205">Method</span></span> |
| [<span data-ttu-id="c6f5f-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="c6f5f-207">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-207">Method</span></span> |
| [<span data-ttu-id="c6f5f-208">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="c6f5f-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="c6f5f-209">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-209">Method</span></span> |
| [<span data-ttu-id="c6f5f-210">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c6f5f-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="c6f5f-211">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-211">Method</span></span> |
| [<span data-ttu-id="c6f5f-212">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="c6f5f-213">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-213">Method</span></span> |
| [<span data-ttu-id="c6f5f-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="c6f5f-215">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-215">Method</span></span> |
| [<span data-ttu-id="c6f5f-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="c6f5f-217">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-217">Method</span></span> |
| [<span data-ttu-id="c6f5f-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="c6f5f-219">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-219">Method</span></span> |
| [<span data-ttu-id="c6f5f-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="c6f5f-221">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-221">Method</span></span> |
| [<span data-ttu-id="c6f5f-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c6f5f-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="c6f5f-223">Método</span><span class="sxs-lookup"><span data-stu-id="c6f5f-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="c6f5f-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-224">Example</span></span>

<span data-ttu-id="c6f5f-225">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="c6f5f-226">Members</span><span class="sxs-lookup"><span data-stu-id="c6f5f-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a><span data-ttu-id="c6f5f-227">anexos: Matriz.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="c6f5f-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

<span data-ttu-id="c6f5f-228">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="c6f5f-229">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-230">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c6f5f-231">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-232">Type</span></span>

*   <span data-ttu-id="c6f5f-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="c6f5f-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-234">Requirements</span></span>

|<span data-ttu-id="c6f5f-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-235">Requirement</span></span>|<span data-ttu-id="c6f5f-236">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-238">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-238">1.0</span></span>|
|[<span data-ttu-id="c6f5f-239">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-240">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-241">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-242">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-243">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-243">Example</span></span>

<span data-ttu-id="c6f5f-244">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="c6f5f-245">cco :[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-245">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-246">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c6f5f-247">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-247">Compose mode only.</span></span>

<span data-ttu-id="c6f5f-248">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-248">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c6f5f-249">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-249">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c6f5f-250">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-250">Get 500 members maximum.</span></span>
- <span data-ttu-id="c6f5f-251">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-251">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-252">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-252">Type</span></span>

*   [<span data-ttu-id="c6f5f-253">Destinatários</span><span class="sxs-lookup"><span data-stu-id="c6f5f-253">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="c6f5f-254">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-254">Requirements</span></span>

|<span data-ttu-id="c6f5f-255">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-255">Requirement</span></span>|<span data-ttu-id="c6f5f-256">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-256">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-257">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-257">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-258">1.1</span><span class="sxs-lookup"><span data-stu-id="c6f5f-258">1.1</span></span>|
|[<span data-ttu-id="c6f5f-259">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-259">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-260">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-260">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-261">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-261">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-262">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-262">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-263">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-263">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-18"></a><span data-ttu-id="c6f5f-264">corpo: [Corpo](/javascript/api/outlook/office.body?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-264">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-265">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-265">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-266">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-266">Type</span></span>

*   [<span data-ttu-id="c6f5f-267">Body</span><span class="sxs-lookup"><span data-stu-id="c6f5f-267">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="c6f5f-268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-268">Requirements</span></span>

|<span data-ttu-id="c6f5f-269">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-269">Requirement</span></span>|<span data-ttu-id="c6f5f-270">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-271">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-272">1.1</span><span class="sxs-lookup"><span data-stu-id="c6f5f-272">1.1</span></span>|
|[<span data-ttu-id="c6f5f-273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-274">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-275">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-276">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-277">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-277">Example</span></span>

<span data-ttu-id="c6f5f-278">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-278">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="c6f5f-279">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-279">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategoriesviewoutlook-js-18"></a><span data-ttu-id="c6f5f-280">Categorias: [categorias](/javascript/api/outlook/office.categories?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-280">categories: [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-281">Obtém um objeto que fornece métodos para gerenciar as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-281">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-282">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-282">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-283">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-283">Type</span></span>

*   [<span data-ttu-id="c6f5f-284">Categories</span><span class="sxs-lookup"><span data-stu-id="c6f5f-284">Categories</span></span>](/javascript/api/outlook/office.categories?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="c6f5f-285">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-285">Requirements</span></span>

|<span data-ttu-id="c6f5f-286">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-286">Requirement</span></span>|<span data-ttu-id="c6f5f-287">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-288">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-289">1,8</span><span class="sxs-lookup"><span data-stu-id="c6f5f-289">1.8</span></span>|
|[<span data-ttu-id="c6f5f-290">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-291">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-292">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-293">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-294">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-294">Example</span></span>

<span data-ttu-id="c6f5f-295">Este exemplo obtém as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-295">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="c6f5f-296">cc : Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-296">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-297">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-297">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c6f5f-298">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-298">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6f5f-299">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="c6f5f-299">Read mode</span></span>

<span data-ttu-id="c6f5f-300">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-300">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="c6f5f-301">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-301">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c6f5f-302">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-302">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="c6f5f-303">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="c6f5f-303">Compose mode</span></span>

<span data-ttu-id="c6f5f-304">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-304">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="c6f5f-305">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-305">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c6f5f-306">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-306">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c6f5f-307">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-307">Get 500 members maximum.</span></span>
- <span data-ttu-id="c6f5f-308">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-308">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c6f5f-309">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-309">Type</span></span>

*   <span data-ttu-id="c6f5f-310">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-310">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-311">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-311">Requirements</span></span>

|<span data-ttu-id="c6f5f-312">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-312">Requirement</span></span>|<span data-ttu-id="c6f5f-313">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-313">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-314">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-314">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-315">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-315">1.0</span></span>|
|[<span data-ttu-id="c6f5f-316">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-316">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-317">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-317">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-318">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-318">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-319">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-319">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="c6f5f-320">(anulável) conversationId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="c6f5f-320">(nullable) conversationId: String</span></span>

<span data-ttu-id="c6f5f-321">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-321">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c6f5f-p109">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c6f5f-p110">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-326">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-326">Type</span></span>

*   <span data-ttu-id="c6f5f-327">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-327">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-328">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-328">Requirements</span></span>

|<span data-ttu-id="c6f5f-329">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-329">Requirement</span></span>|<span data-ttu-id="c6f5f-330">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-330">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-331">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-331">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-332">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-332">1.0</span></span>|
|[<span data-ttu-id="c6f5f-333">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-333">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-334">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-334">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-335">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-335">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-336">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-336">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-337">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-337">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="c6f5f-338">dateTimeCreated: Data</span><span class="sxs-lookup"><span data-stu-id="c6f5f-338">dateTimeCreated: Date</span></span>

<span data-ttu-id="c6f5f-p111">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-341">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-341">Type</span></span>

*   <span data-ttu-id="c6f5f-342">Data</span><span class="sxs-lookup"><span data-stu-id="c6f5f-342">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-343">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-343">Requirements</span></span>

|<span data-ttu-id="c6f5f-344">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-344">Requirement</span></span>|<span data-ttu-id="c6f5f-345">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-346">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-347">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-347">1.0</span></span>|
|[<span data-ttu-id="c6f5f-348">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-349">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-350">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-351">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-352">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-352">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="c6f5f-353">dateTimeModified: Data</span><span class="sxs-lookup"><span data-stu-id="c6f5f-353">dateTimeModified: Date</span></span>

<span data-ttu-id="c6f5f-p112">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-356">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-356">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-357">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-357">Type</span></span>

*   <span data-ttu-id="c6f5f-358">Data</span><span class="sxs-lookup"><span data-stu-id="c6f5f-358">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-359">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-359">Requirements</span></span>

|<span data-ttu-id="c6f5f-360">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-360">Requirement</span></span>|<span data-ttu-id="c6f5f-361">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-362">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-363">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-363">1.0</span></span>|
|[<span data-ttu-id="c6f5f-364">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-365">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-366">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-367">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-367">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-368">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-368">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a><span data-ttu-id="c6f5f-369">fim: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-369">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-370">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-370">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c6f5f-p113">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6f5f-373">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="c6f5f-373">Read mode</span></span>

<span data-ttu-id="c6f5f-374">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-374">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="c6f5f-375">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="c6f5f-375">Compose mode</span></span>

<span data-ttu-id="c6f5f-376">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-376">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c6f5f-377">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-377">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="c6f5f-378">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-378">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="c6f5f-379">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-379">Type</span></span>

*   <span data-ttu-id="c6f5f-380">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-380">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-381">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-381">Requirements</span></span>

|<span data-ttu-id="c6f5f-382">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-382">Requirement</span></span>|<span data-ttu-id="c6f5f-383">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-383">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-384">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-384">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-385">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-385">1.0</span></span>|
|[<span data-ttu-id="c6f5f-386">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-386">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-387">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-387">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-388">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-388">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-389">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-389">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocationviewoutlook-js-18"></a><span data-ttu-id="c6f5f-390">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-390">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-391">Obtém ou define os locais de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-391">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6f5f-392">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="c6f5f-392">Read mode</span></span>

<span data-ttu-id="c6f5f-393">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) que permite que você obtenha o conjunto de locais (cada um representado por um objeto [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8) ) associado ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-393">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c6f5f-394">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="c6f5f-394">Compose mode</span></span>

<span data-ttu-id="c6f5f-395">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) que fornece métodos para obter, remover ou adicionar locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-395">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-396">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-396">Type</span></span>

*   [<span data-ttu-id="c6f5f-397">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="c6f5f-397">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="c6f5f-398">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-398">Requirements</span></span>

|<span data-ttu-id="c6f5f-399">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-399">Requirement</span></span>|<span data-ttu-id="c6f5f-400">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-400">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-401">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-401">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-402">1,8</span><span class="sxs-lookup"><span data-stu-id="c6f5f-402">1.8</span></span>|
|[<span data-ttu-id="c6f5f-403">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-403">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-404">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-404">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-405">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-405">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-406">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-406">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-407">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-407">Example</span></span>

<span data-ttu-id="c6f5f-408">O exemplo a seguir obtém os locais atuais associados ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-408">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18fromjavascriptapioutlookofficefromviewoutlook-js-18"></a><span data-ttu-id="c6f5f-409">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[de](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-409">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-410">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-410">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="c6f5f-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-413">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-413">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6f5f-414">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="c6f5f-414">Read mode</span></span>

<span data-ttu-id="c6f5f-415">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-415">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="c6f5f-416">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="c6f5f-416">Compose mode</span></span>

<span data-ttu-id="c6f5f-417">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-417">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c6f5f-418">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-418">Type</span></span>

*   <span data-ttu-id="c6f5f-419">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [de](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-419">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-420">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-420">Requirements</span></span>

|<span data-ttu-id="c6f5f-421">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-421">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c6f5f-422">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-423">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-423">1.0</span></span>|<span data-ttu-id="c6f5f-424">1.7</span><span class="sxs-lookup"><span data-stu-id="c6f5f-424">1.7</span></span>|
|[<span data-ttu-id="c6f5f-425">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-426">ReadItem</span></span>|<span data-ttu-id="c6f5f-427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-427">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6f5f-428">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-429">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-429">Read</span></span>|<span data-ttu-id="c6f5f-430">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-430">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheadersviewoutlook-js-18"></a><span data-ttu-id="c6f5f-431">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-431">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-432">Obtém ou define cabeçalhos de Internet personalizados em uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-432">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="c6f5f-433">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-433">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-434">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-434">Type</span></span>

*   [<span data-ttu-id="c6f5f-435">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="c6f5f-435">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="c6f5f-436">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-436">Requirements</span></span>

|<span data-ttu-id="c6f5f-437">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-437">Requirement</span></span>|<span data-ttu-id="c6f5f-438">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-438">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-439">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-440">1,8</span><span class="sxs-lookup"><span data-stu-id="c6f5f-440">1.8</span></span>|
|[<span data-ttu-id="c6f5f-441">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-441">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-442">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-443">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-443">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-444">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-444">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-445">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-445">Example</span></span>

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

#### <a name="internetmessageid-string"></a><span data-ttu-id="c6f5f-446">internetMessageId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="c6f5f-446">internetMessageId: String</span></span>

<span data-ttu-id="c6f5f-p116">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-449">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-449">Type</span></span>

*   <span data-ttu-id="c6f5f-450">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-450">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-451">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-451">Requirements</span></span>

|<span data-ttu-id="c6f5f-452">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-452">Requirement</span></span>|<span data-ttu-id="c6f5f-453">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-453">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-454">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-454">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-455">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-455">1.0</span></span>|
|[<span data-ttu-id="c6f5f-456">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-456">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-457">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-457">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-458">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-458">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-459">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-459">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-460">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-460">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="c6f5f-461">itemClass: Cadeia de caracteres </span><span class="sxs-lookup"><span data-stu-id="c6f5f-461">itemClass: String</span></span>

<span data-ttu-id="c6f5f-p117">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c6f5f-p118">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="c6f5f-466">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-466">Type</span></span>|<span data-ttu-id="c6f5f-467">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-467">Description</span></span>|<span data-ttu-id="c6f5f-468">classe de item</span><span class="sxs-lookup"><span data-stu-id="c6f5f-468">item class</span></span>|
|---|---|---|
|<span data-ttu-id="c6f5f-469">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="c6f5f-469">Appointment items</span></span>|<span data-ttu-id="c6f5f-470">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-470">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="c6f5f-471">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-471">Message items</span></span>|<span data-ttu-id="c6f5f-472">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-472">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="c6f5f-473">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-473">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-474">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-474">Type</span></span>

*   <span data-ttu-id="c6f5f-475">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-475">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-476">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-476">Requirements</span></span>

|<span data-ttu-id="c6f5f-477">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-477">Requirement</span></span>|<span data-ttu-id="c6f5f-478">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-478">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-479">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-479">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-480">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-480">1.0</span></span>|
|[<span data-ttu-id="c6f5f-481">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-481">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-482">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-482">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-483">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-483">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-484">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-484">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-485">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-485">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c6f5f-486">(anulável) itemId: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c6f5f-486">(nullable) itemId: String</span></span>

<span data-ttu-id="c6f5f-p119">Obtém o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-489">O identificador retornado pela propriedade `itemId` é o mesmo que o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-489">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="c6f5f-490">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-490">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c6f5f-491">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-491">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c6f5f-492">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-492">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="c6f5f-p121">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-495">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-495">Type</span></span>

*   <span data-ttu-id="c6f5f-496">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-496">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-497">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-497">Requirements</span></span>

|<span data-ttu-id="c6f5f-498">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-498">Requirement</span></span>|<span data-ttu-id="c6f5f-499">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-499">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-500">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-500">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-501">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-501">1.0</span></span>|
|[<span data-ttu-id="c6f5f-502">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-502">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-503">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-503">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-504">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-504">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-505">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-505">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-506">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-506">Example</span></span>

<span data-ttu-id="c6f5f-p122">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-18"></a><span data-ttu-id="c6f5f-509">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-509">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-510">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-510">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c6f5f-511">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-511">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-512">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-512">Type</span></span>

*   [<span data-ttu-id="c6f5f-513">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="c6f5f-513">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="c6f5f-514">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-514">Requirements</span></span>

|<span data-ttu-id="c6f5f-515">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-515">Requirement</span></span>|<span data-ttu-id="c6f5f-516">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-516">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-517">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-517">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-518">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-518">1.0</span></span>|
|[<span data-ttu-id="c6f5f-519">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-519">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-520">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-520">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-521">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-521">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-522">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-522">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-523">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-523">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-18"></a><span data-ttu-id="c6f5f-524">Local: Cadeia de caracteres[Local](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-524">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-525">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-525">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6f5f-526">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="c6f5f-526">Read mode</span></span>

<span data-ttu-id="c6f5f-527">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-527">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="c6f5f-528">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="c6f5f-528">Compose mode</span></span>

<span data-ttu-id="c6f5f-529">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-529">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c6f5f-530">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-530">Type</span></span>

*   <span data-ttu-id="c6f5f-531">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-531">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-532">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-532">Requirements</span></span>

|<span data-ttu-id="c6f5f-533">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-533">Requirement</span></span>|<span data-ttu-id="c6f5f-534">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-534">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-535">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-535">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-536">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-536">1.0</span></span>|
|[<span data-ttu-id="c6f5f-537">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-537">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-538">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-538">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-539">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-539">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-540">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-540">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c6f5f-541">normalizedSubject: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c6f5f-541">normalizedSubject: String</span></span>

<span data-ttu-id="c6f5f-p123">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c6f5f-p124">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-546">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-546">Type</span></span>

*   <span data-ttu-id="c6f5f-547">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-547">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-548">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-548">Requirements</span></span>

|<span data-ttu-id="c6f5f-549">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-549">Requirement</span></span>|<span data-ttu-id="c6f5f-550">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-551">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-552">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-552">1.0</span></span>|
|[<span data-ttu-id="c6f5f-553">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-553">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-554">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-555">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-555">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-556">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-556">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-557">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-557">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-18"></a><span data-ttu-id="c6f5f-558">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-558">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-559">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-559">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-560">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-560">Type</span></span>

*   [<span data-ttu-id="c6f5f-561">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c6f5f-561">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="c6f5f-562">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-562">Requirements</span></span>

|<span data-ttu-id="c6f5f-563">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-563">Requirement</span></span>|<span data-ttu-id="c6f5f-564">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-565">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-566">1.3</span><span class="sxs-lookup"><span data-stu-id="c6f5f-566">1.3</span></span>|
|[<span data-ttu-id="c6f5f-567">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-567">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-568">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-568">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-569">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-569">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-570">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-570">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-571">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-571">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="c6f5f-572">optionalAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-572">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-573">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-573">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c6f5f-574">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-574">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6f5f-575">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="c6f5f-575">Read mode</span></span>

<span data-ttu-id="c6f5f-576">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-576">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="c6f5f-577">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-577">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c6f5f-578">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-578">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="c6f5f-579">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="c6f5f-579">Compose mode</span></span>

<span data-ttu-id="c6f5f-580">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-580">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="c6f5f-581">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-581">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c6f5f-582">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-582">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c6f5f-583">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-583">Get 500 members maximum.</span></span>
- <span data-ttu-id="c6f5f-584">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-584">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c6f5f-585">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-585">Type</span></span>

*   <span data-ttu-id="c6f5f-586">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-586">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-587">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-587">Requirements</span></span>

|<span data-ttu-id="c6f5f-588">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-588">Requirement</span></span>|<span data-ttu-id="c6f5f-589">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-589">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-590">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-591">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-591">1.0</span></span>|
|[<span data-ttu-id="c6f5f-592">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-592">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-593">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-593">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-594">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-594">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-595">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-595">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18organizerjavascriptapioutlookofficeorganizerviewoutlook-js-18"></a><span data-ttu-id="c6f5f-596">organizador: [](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c6f5f-596">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-597">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-597">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6f5f-598">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="c6f5f-598">Read mode</span></span>

<span data-ttu-id="c6f5f-599">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-599">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="c6f5f-600">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="c6f5f-600">Compose mode</span></span>

<span data-ttu-id="c6f5f-601">A `organizer` propriedade retorna um objeto [organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-601">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="c6f5f-602">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-602">Type</span></span>

*   <span data-ttu-id="c6f5f-603">[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [Organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c6f5f-603">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-604">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-604">Requirements</span></span>

|<span data-ttu-id="c6f5f-605">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-605">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c6f5f-606">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-607">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-607">1.0</span></span>|<span data-ttu-id="c6f5f-608">1.7</span><span class="sxs-lookup"><span data-stu-id="c6f5f-608">1.7</span></span>|
|[<span data-ttu-id="c6f5f-609">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-610">ReadItem</span></span>|<span data-ttu-id="c6f5f-611">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-611">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6f5f-612">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-613">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-613">Read</span></span>|<span data-ttu-id="c6f5f-614">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-614">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-18"></a><span data-ttu-id="c6f5f-615">(anulável) recorrência: [recorrência](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-615">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-616">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-616">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="c6f5f-617">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-617">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="c6f5f-618">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-618">Read and compose modes for appointment items.</span></span> <span data-ttu-id="c6f5f-619">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-619">Read mode for meeting request items.</span></span>

<span data-ttu-id="c6f5f-620">A `recurrence` propriedade retorna um objeto de [recorrência](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-620">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="c6f5f-621">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-621">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="c6f5f-622">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-622">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="c6f5f-623">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-623">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="c6f5f-624">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-624">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6f5f-625">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="c6f5f-625">Read mode</span></span>

<span data-ttu-id="c6f5f-626">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-626">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object that represents the appointment recurrence.</span></span> <span data-ttu-id="c6f5f-627">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-627">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="c6f5f-628">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="c6f5f-628">Compose mode</span></span>

<span data-ttu-id="c6f5f-629">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-629">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="c6f5f-630">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-630">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="c6f5f-631">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-631">Type</span></span>

* [<span data-ttu-id="c6f5f-632">Recorrência</span><span class="sxs-lookup"><span data-stu-id="c6f5f-632">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)

|<span data-ttu-id="c6f5f-633">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-633">Requirement</span></span>|<span data-ttu-id="c6f5f-634">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-635">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-636">1.7</span><span class="sxs-lookup"><span data-stu-id="c6f5f-636">1.7</span></span>|
|[<span data-ttu-id="c6f5f-637">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-638">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-639">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-640">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-640">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="c6f5f-641">requiredAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-641">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-642">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-642">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c6f5f-643">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-643">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6f5f-644">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="c6f5f-644">Read mode</span></span>

<span data-ttu-id="c6f5f-645">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-645">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="c6f5f-646">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-646">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c6f5f-647">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-647">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="c6f5f-648">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="c6f5f-648">Compose mode</span></span>

<span data-ttu-id="c6f5f-649">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-649">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="c6f5f-650">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-650">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c6f5f-651">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-651">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c6f5f-652">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-652">Get 500 members maximum.</span></span>
- <span data-ttu-id="c6f5f-653">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-653">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="c6f5f-654">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-654">Type</span></span>

*   <span data-ttu-id="c6f5f-655">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-655">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-656">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-656">Requirements</span></span>

|<span data-ttu-id="c6f5f-657">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-657">Requirement</span></span>|<span data-ttu-id="c6f5f-658">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-658">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-659">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-659">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-660">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-660">1.0</span></span>|
|[<span data-ttu-id="c6f5f-661">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-661">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-662">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-662">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-663">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-663">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-664">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-664">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18"></a><span data-ttu-id="c6f5f-665">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-665">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-p135">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c6f5f-p136">As propriedades [`from`](#from-emailaddressdetailsfrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-670">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-670">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-671">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-671">Type</span></span>

*   [<span data-ttu-id="c6f5f-672">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c6f5f-672">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="c6f5f-673">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-673">Requirements</span></span>

|<span data-ttu-id="c6f5f-674">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-674">Requirement</span></span>|<span data-ttu-id="c6f5f-675">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-675">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-676">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-676">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-677">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-677">1.0</span></span>|
|[<span data-ttu-id="c6f5f-678">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-678">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-679">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-679">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-680">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-680">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-681">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-681">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-682">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-682">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="c6f5f-683">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-683">(nullable) seriesId: String</span></span>

<span data-ttu-id="c6f5f-684">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-684">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="c6f5f-685">No Outlook na Web e clientes de desktop, o `seriesId` retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-685">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="c6f5f-686">No entanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-686">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-687">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-687">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c6f5f-688">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-688">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="c6f5f-689">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-689">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c6f5f-690">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-690">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="c6f5f-691">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-691">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="c6f5f-692">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-692">Type</span></span>

* <span data-ttu-id="c6f5f-693">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-693">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-694">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-694">Requirements</span></span>

|<span data-ttu-id="c6f5f-695">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-695">Requirement</span></span>|<span data-ttu-id="c6f5f-696">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-697">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-698">1.7</span><span class="sxs-lookup"><span data-stu-id="c6f5f-698">1.7</span></span>|
|[<span data-ttu-id="c6f5f-699">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-700">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-701">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-702">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-702">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-703">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-703">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a><span data-ttu-id="c6f5f-704">início: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-704">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-705">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-705">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c6f5f-p139">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6f5f-708">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="c6f5f-708">Read mode</span></span>

<span data-ttu-id="c6f5f-709">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-709">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="c6f5f-710">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="c6f5f-710">Compose mode</span></span>

<span data-ttu-id="c6f5f-711">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-711">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c6f5f-712">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-712">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="c6f5f-713">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-713">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="c6f5f-714">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-714">Type</span></span>

*   <span data-ttu-id="c6f5f-715">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-715">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-716">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-716">Requirements</span></span>

|<span data-ttu-id="c6f5f-717">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-717">Requirement</span></span>|<span data-ttu-id="c6f5f-718">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-719">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-719">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-720">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-720">1.0</span></span>|
|[<span data-ttu-id="c6f5f-721">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-721">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-722">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-723">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-723">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-724">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-724">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-18"></a><span data-ttu-id="c6f5f-725">Assunto: Cadeia de caracteres|[Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-725">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-726">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-726">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c6f5f-727">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-727">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6f5f-728">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="c6f5f-728">Read mode</span></span>

<span data-ttu-id="c6f5f-p140">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="c6f5f-731">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-731">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="c6f5f-732">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="c6f5f-732">Compose mode</span></span>
<span data-ttu-id="c6f5f-733">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-733">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="c6f5f-734">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-734">Type</span></span>

*   <span data-ttu-id="c6f5f-735">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-735">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-736">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-736">Requirements</span></span>

|<span data-ttu-id="c6f5f-737">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-737">Requirement</span></span>|<span data-ttu-id="c6f5f-738">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-738">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-739">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-739">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-740">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-740">1.0</span></span>|
|[<span data-ttu-id="c6f5f-741">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-741">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-742">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-742">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-743">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-743">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-744">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-744">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="c6f5f-745">para: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-745">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-746">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-746">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c6f5f-747">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-747">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c6f5f-748">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="c6f5f-748">Read mode</span></span>

<span data-ttu-id="c6f5f-749">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-749">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="c6f5f-750">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-750">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c6f5f-751">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-751">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="c6f5f-752">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="c6f5f-752">Compose mode</span></span>

<span data-ttu-id="c6f5f-753">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-753">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="c6f5f-754">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-754">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c6f5f-755">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-755">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c6f5f-756">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-756">Get 500 members maximum.</span></span>
- <span data-ttu-id="c6f5f-757">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-757">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c6f5f-758">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-758">Type</span></span>

*   <span data-ttu-id="c6f5f-759">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-759">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-760">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-760">Requirements</span></span>

|<span data-ttu-id="c6f5f-761">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-761">Requirement</span></span>|<span data-ttu-id="c6f5f-762">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-762">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-763">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-763">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-764">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-764">1.0</span></span>|
|[<span data-ttu-id="c6f5f-765">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-765">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-766">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-766">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-767">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-767">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-768">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-768">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="c6f5f-769">Métodos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-769">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c6f5f-770">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c6f5f-770">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c6f5f-771">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-771">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c6f5f-772">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-772">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c6f5f-773">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-773">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-774">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-774">Parameters</span></span>
|<span data-ttu-id="c6f5f-775">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-775">Name</span></span>|<span data-ttu-id="c6f5f-776">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-776">Type</span></span>|<span data-ttu-id="c6f5f-777">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-777">Attributes</span></span>|<span data-ttu-id="c6f5f-778">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-778">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="c6f5f-779">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-779">String</span></span>||<span data-ttu-id="c6f5f-p144">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c6f5f-782">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-782">String</span></span>||<span data-ttu-id="c6f5f-p145">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c6f5f-785">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-785">Object</span></span>|<span data-ttu-id="c6f5f-786">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-786">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-787">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-787">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c6f5f-788">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-788">Object</span></span>|<span data-ttu-id="c6f5f-789">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-789">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-790">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-790">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="c6f5f-791">Booliano</span><span class="sxs-lookup"><span data-stu-id="c6f5f-791">Boolean</span></span>|<span data-ttu-id="c6f5f-792">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-792">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-793">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-793">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-794">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-794">function</span></span>|<span data-ttu-id="c6f5f-795">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-795">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-796">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-796">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c6f5f-797">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-797">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c6f5f-798">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-798">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c6f5f-799">Erros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-799">Errors</span></span>

|<span data-ttu-id="c6f5f-800">Código de erro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-800">Error code</span></span>|<span data-ttu-id="c6f5f-801">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-801">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="c6f5f-802">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-802">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="c6f5f-803">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-803">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c6f5f-804">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-804">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-805">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-805">Requirements</span></span>

|<span data-ttu-id="c6f5f-806">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-806">Requirement</span></span>|<span data-ttu-id="c6f5f-807">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-808">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-809">1.1</span><span class="sxs-lookup"><span data-stu-id="c6f5f-809">1.1</span></span>|
|[<span data-ttu-id="c6f5f-810">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-811">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-811">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6f5f-812">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-813">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-813">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c6f5f-814">Exemplos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-814">Examples</span></span>

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

<span data-ttu-id="c6f5f-815">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-815">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="c6f5f-816">addFileAttachmentFromBase64Async (base64file, AttachmentName, [Options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c6f5f-816">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c6f5f-817">Adiciona um arquivo da codificação Base64 a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-817">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c6f5f-818">O `addFileAttachmentFromBase64Async` método carrega o arquivo da codificação Base64 e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-818">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="c6f5f-819">Esse método retorna o identificador de anexo no objeto AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-819">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="c6f5f-820">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-820">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-821">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-821">Parameters</span></span>

|<span data-ttu-id="c6f5f-822">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-822">Name</span></span>|<span data-ttu-id="c6f5f-823">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-823">Type</span></span>|<span data-ttu-id="c6f5f-824">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-824">Attributes</span></span>|<span data-ttu-id="c6f5f-825">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-825">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="c6f5f-826">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-826">String</span></span>||<span data-ttu-id="c6f5f-827">O conteúdo codificado em Base64 de uma imagem ou arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-827">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="c6f5f-828">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-828">String</span></span>||<span data-ttu-id="c6f5f-p147">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c6f5f-831">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-831">Object</span></span>|<span data-ttu-id="c6f5f-832">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-832">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-833">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-833">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c6f5f-834">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-834">Object</span></span>|<span data-ttu-id="c6f5f-835">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-835">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-836">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-836">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="c6f5f-837">Booliano</span><span class="sxs-lookup"><span data-stu-id="c6f5f-837">Boolean</span></span>|<span data-ttu-id="c6f5f-838">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-838">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-839">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-839">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-840">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-840">function</span></span>|<span data-ttu-id="c6f5f-841">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-841">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-842">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-842">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c6f5f-843">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-843">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c6f5f-844">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-844">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c6f5f-845">Erros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-845">Errors</span></span>

|<span data-ttu-id="c6f5f-846">Código de erro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-846">Error code</span></span>|<span data-ttu-id="c6f5f-847">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-847">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="c6f5f-848">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-848">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="c6f5f-849">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-849">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c6f5f-850">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-850">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-851">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-851">Requirements</span></span>

|<span data-ttu-id="c6f5f-852">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-852">Requirement</span></span>|<span data-ttu-id="c6f5f-853">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-853">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-854">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-854">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-855">1,8</span><span class="sxs-lookup"><span data-stu-id="c6f5f-855">1.8</span></span>|
|[<span data-ttu-id="c6f5f-856">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-856">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-857">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-857">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6f5f-858">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-858">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-859">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-859">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c6f5f-860">Exemplos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-860">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c6f5f-861">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c6f5f-861">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c6f5f-862">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-862">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c6f5f-863">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-863">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-864">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-864">Parameters</span></span>

| <span data-ttu-id="c6f5f-865">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-865">Name</span></span> | <span data-ttu-id="c6f5f-866">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-866">Type</span></span> | <span data-ttu-id="c6f5f-867">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-867">Attributes</span></span> | <span data-ttu-id="c6f5f-868">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-868">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c6f5f-869">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c6f5f-869">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c6f5f-870">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-870">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c6f5f-871">Função</span><span class="sxs-lookup"><span data-stu-id="c6f5f-871">Function</span></span> || <span data-ttu-id="c6f5f-p148">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c6f5f-875">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-875">Object</span></span> | <span data-ttu-id="c6f5f-876">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-876">&lt;optional&gt;</span></span> | <span data-ttu-id="c6f5f-877">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-877">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c6f5f-878">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-878">Object</span></span> | <span data-ttu-id="c6f5f-879">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-879">&lt;optional&gt;</span></span> | <span data-ttu-id="c6f5f-880">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-880">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c6f5f-881">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-881">function</span></span>| <span data-ttu-id="c6f5f-882">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-882">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-883">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-884">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-884">Requirements</span></span>

|<span data-ttu-id="c6f5f-885">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-885">Requirement</span></span>| <span data-ttu-id="c6f5f-886">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-887">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6f5f-888">1.7</span><span class="sxs-lookup"><span data-stu-id="c6f5f-888">1.7</span></span> |
|[<span data-ttu-id="c6f5f-889">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6f5f-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-890">ReadItem</span></span> |
|[<span data-ttu-id="c6f5f-891">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6f5f-892">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-892">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="c6f5f-893">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-893">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c6f5f-894">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c6f5f-894">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c6f5f-895">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-895">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c6f5f-p149">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c6f5f-899">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-899">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c6f5f-900">Se o Suplemento do Office estiver em execução no Outlook na Web, o método `addItemAttachmentAsync` pode anexar itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-900">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-901">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-901">Parameters</span></span>

|<span data-ttu-id="c6f5f-902">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-902">Name</span></span>|<span data-ttu-id="c6f5f-903">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-903">Type</span></span>|<span data-ttu-id="c6f5f-904">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-904">Attributes</span></span>|<span data-ttu-id="c6f5f-905">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-905">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="c6f5f-906">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-906">String</span></span>||<span data-ttu-id="c6f5f-p150">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c6f5f-909">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c6f5f-909">String</span></span>||<span data-ttu-id="c6f5f-910">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-910">The subject of the item to be attached.</span></span> <span data-ttu-id="c6f5f-911">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-911">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c6f5f-912">Object</span><span class="sxs-lookup"><span data-stu-id="c6f5f-912">Object</span></span>|<span data-ttu-id="c6f5f-913">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-913">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-914">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-914">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c6f5f-915">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-915">Object</span></span>|<span data-ttu-id="c6f5f-916">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-916">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-917">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-917">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-918">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-918">function</span></span>|<span data-ttu-id="c6f5f-919">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-919">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-920">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-920">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c6f5f-921">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-921">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c6f5f-922">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-922">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c6f5f-923">Erros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-923">Errors</span></span>

|<span data-ttu-id="c6f5f-924">Código de erro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-924">Error code</span></span>|<span data-ttu-id="c6f5f-925">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-925">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c6f5f-926">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-926">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-927">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-927">Requirements</span></span>

|<span data-ttu-id="c6f5f-928">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-928">Requirement</span></span>|<span data-ttu-id="c6f5f-929">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-929">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-930">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-930">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-931">1.1</span><span class="sxs-lookup"><span data-stu-id="c6f5f-931">1.1</span></span>|
|[<span data-ttu-id="c6f5f-932">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-932">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-933">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-933">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6f5f-934">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-934">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-935">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-935">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-936">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-936">Example</span></span>

<span data-ttu-id="c6f5f-937">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-937">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="c6f5f-938">close()</span><span class="sxs-lookup"><span data-stu-id="c6f5f-938">close()</span></span>

<span data-ttu-id="c6f5f-939">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-939">Closes the current item that is being composed.</span></span>

<span data-ttu-id="c6f5f-p152">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-942">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-942">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="c6f5f-943">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-943">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-944">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-944">Requirements</span></span>

|<span data-ttu-id="c6f5f-945">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-945">Requirement</span></span>|<span data-ttu-id="c6f5f-946">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-947">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-948">1.3</span><span class="sxs-lookup"><span data-stu-id="c6f5f-948">1.3</span></span>|
|[<span data-ttu-id="c6f5f-949">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-950">Restrito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-950">Restricted</span></span>|
|[<span data-ttu-id="c6f5f-951">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-952">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-952">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="c6f5f-953">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="c6f5f-953">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="c6f5f-954">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-954">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-955">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-955">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c6f5f-956">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-956">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c6f5f-957">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-957">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c6f5f-p153">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-961">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-961">Parameters</span></span>

|<span data-ttu-id="c6f5f-962">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-962">Name</span></span>|<span data-ttu-id="c6f5f-963">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-963">Type</span></span>|<span data-ttu-id="c6f5f-964">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-964">Attributes</span></span>|<span data-ttu-id="c6f5f-965">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-965">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c6f5f-966">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c6f5f-966">String &#124; Object</span></span>||<span data-ttu-id="c6f5f-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c6f5f-969">**OU**</span><span class="sxs-lookup"><span data-stu-id="c6f5f-969">**OR**</span></span><br/><span data-ttu-id="c6f5f-p155">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c6f5f-972">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-972">String</span></span>|<span data-ttu-id="c6f5f-973">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-973">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-p156">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c6f5f-976">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-976">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c6f5f-977">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-977">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-978">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-978">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c6f5f-979">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-979">String</span></span>||<span data-ttu-id="c6f5f-p157">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c6f5f-982">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c6f5f-982">String</span></span>||<span data-ttu-id="c6f5f-983">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-983">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c6f5f-984">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c6f5f-984">String</span></span>||<span data-ttu-id="c6f5f-p158">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c6f5f-987">Booliano</span><span class="sxs-lookup"><span data-stu-id="c6f5f-987">Boolean</span></span>||<span data-ttu-id="c6f5f-p159">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c6f5f-990">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-990">String</span></span>||<span data-ttu-id="c6f5f-p160">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-994">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-994">function</span></span>|<span data-ttu-id="c6f5f-995">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-995">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-996">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-996">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-997">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-997">Requirements</span></span>

|<span data-ttu-id="c6f5f-998">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-998">Requirement</span></span>|<span data-ttu-id="c6f5f-999">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-999">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1000">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1000">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1001">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1001">1.0</span></span>|
|[<span data-ttu-id="c6f5f-1002">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1002">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1003">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1003">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1004">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1004">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1005">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1005">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c6f5f-1006">Exemplos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1006">Examples</span></span>

<span data-ttu-id="c6f5f-1007">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1007">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c6f5f-1008">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1008">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c6f5f-1009">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1009">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c6f5f-1010">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1010">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c6f5f-1011">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1011">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c6f5f-1012">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1012">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="c6f5f-1013">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1013">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="c6f5f-1014">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1014">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-1015">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1015">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c6f5f-1016">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1016">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c6f5f-1017">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1017">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c6f5f-p161">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1021">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1021">Parameters</span></span>

|<span data-ttu-id="c6f5f-1022">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1022">Name</span></span>|<span data-ttu-id="c6f5f-1023">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1023">Type</span></span>|<span data-ttu-id="c6f5f-1024">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1024">Attributes</span></span>|<span data-ttu-id="c6f5f-1025">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1025">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c6f5f-1026">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1026">String &#124; Object</span></span>||<span data-ttu-id="c6f5f-p162">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c6f5f-1029">**OU**</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1029">**OR**</span></span><br/><span data-ttu-id="c6f5f-p163">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c6f5f-1032">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1032">String</span></span>|<span data-ttu-id="c6f5f-1033">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1033">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-p164">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c6f5f-1036">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1036">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c6f5f-1037">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1038">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1038">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c6f5f-1039">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1039">String</span></span>||<span data-ttu-id="c6f5f-p165">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c6f5f-1042">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1042">String</span></span>||<span data-ttu-id="c6f5f-1043">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1043">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c6f5f-1044">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1044">String</span></span>||<span data-ttu-id="c6f5f-p166">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c6f5f-1047">Booliano</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1047">Boolean</span></span>||<span data-ttu-id="c6f5f-p167">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c6f5f-1050">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1050">String</span></span>||<span data-ttu-id="c6f5f-p168">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-1054">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1054">function</span></span>|<span data-ttu-id="c6f5f-1055">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1055">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1056">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1056">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1057">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1057">Requirements</span></span>

|<span data-ttu-id="c6f5f-1058">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1058">Requirement</span></span>|<span data-ttu-id="c6f5f-1059">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1059">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1060">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1060">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1061">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1061">1.0</span></span>|
|[<span data-ttu-id="c6f5f-1062">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1062">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1063">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1063">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1064">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1064">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1065">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1065">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c6f5f-1066">Exemplos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1066">Examples</span></span>

<span data-ttu-id="c6f5f-1067">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1067">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c6f5f-1068">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1068">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c6f5f-1069">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1069">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c6f5f-1070">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1070">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c6f5f-1071">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1071">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c6f5f-1072">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1072">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="c6f5f-1073">getAllInternetHeadersAsync ([opções], [callback])</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1073">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="c6f5f-1074">Obtém todos os cabeçalhos de Internet da mensagem como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1074">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="c6f5f-1075">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1075">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1076">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1076">Parameters</span></span>

|<span data-ttu-id="c6f5f-1077">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1077">Name</span></span>|<span data-ttu-id="c6f5f-1078">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1078">Type</span></span>|<span data-ttu-id="c6f5f-1079">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1079">Attributes</span></span>|<span data-ttu-id="c6f5f-1080">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1080">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c6f5f-1081">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1081">Object</span></span>|<span data-ttu-id="c6f5f-1082">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1082">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1083">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1083">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c6f5f-1084">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1084">Object</span></span>|<span data-ttu-id="c6f5f-1085">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1085">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1086">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1086">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-1087">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1087">function</span></span>|<span data-ttu-id="c6f5f-1088">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1088">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1089">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="c6f5f-1090">Com êxito, os dados de cabeçalhos de Internet são fornecidos na propriedade asyncResult. Value como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1090">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="c6f5f-1091">Consulte [RFC 2183](https://tools.ietf.org/html/rfc2183) para obter as informações de formatação do valor de cadeia de caracteres retornado.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1091">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="c6f5f-1092">Se a chamada falhar, a propriedade asyncResult. Error conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1092">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1093">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1093">Requirements</span></span>

|<span data-ttu-id="c6f5f-1094">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1094">Requirement</span></span>|<span data-ttu-id="c6f5f-1095">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1096">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1097">1,8</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1097">1.8</span></span>|
|[<span data-ttu-id="c6f5f-1098">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1098">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1099">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1100">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1100">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1101">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1101">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6f5f-1102">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1102">Returns:</span></span>

<span data-ttu-id="c6f5f-1103">A Internet cabeçalhos dados como uma cadeia de caracteres formatada de acordo com a [RFC 2183](https://tools.ietf.org/html/rfc2183).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1103">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="c6f5f-1104">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1104">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c6f5f-1105">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1105">Example</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontentviewoutlook-js-18"></a><span data-ttu-id="c6f5f-1106">getAttachmentContentAsync (attachmentid, [opções], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1106">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span></span>

<span data-ttu-id="c6f5f-1107">Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um `AttachmentContent` objeto.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1107">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="c6f5f-1108">O `getAttachmentContentAsync` método obtém o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1108">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="c6f5f-1109">Como prática recomendada, você deve usar o identificador para recuperar um anexo na mesma sessão em que o attachmentIds foi recuperado com a `getAttachmentsAsync` chamada ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1109">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="c6f5f-1110">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1110">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="c6f5f-1111">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1111">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1112">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1112">Parameters</span></span>

|<span data-ttu-id="c6f5f-1113">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1113">Name</span></span>|<span data-ttu-id="c6f5f-1114">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1114">Type</span></span>|<span data-ttu-id="c6f5f-1115">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1115">Attributes</span></span>|<span data-ttu-id="c6f5f-1116">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1116">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="c6f5f-1117">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1117">String</span></span>||<span data-ttu-id="c6f5f-1118">O identificador do anexo que você deseja obter.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1118">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="c6f5f-1119">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1119">Object</span></span>|<span data-ttu-id="c6f5f-1120">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1120">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1121">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1121">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c6f5f-1122">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1122">Object</span></span>|<span data-ttu-id="c6f5f-1123">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1123">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1124">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1124">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-1125">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1125">function</span></span>|<span data-ttu-id="c6f5f-1126">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1126">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1127">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1128">Requirements</span></span>

|<span data-ttu-id="c6f5f-1129">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1129">Requirement</span></span>|<span data-ttu-id="c6f5f-1130">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1130">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1132">1,8</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1132">1.8</span></span>|
|[<span data-ttu-id="c6f5f-1133">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1133">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1134">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1135">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1136">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1136">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6f5f-1137">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1137">Returns:</span></span>

<span data-ttu-id="c6f5f-1138">Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1138">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="c6f5f-1139">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1139">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a><span data-ttu-id="c6f5f-1140">getAttachmentsAsync ([Options], [callback]) → array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="c6f5f-1140">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

<span data-ttu-id="c6f5f-1141">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1141">Gets the item's attachments as an array.</span></span> <span data-ttu-id="c6f5f-1142">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1142">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1143">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1143">Parameters</span></span>

|<span data-ttu-id="c6f5f-1144">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1144">Name</span></span>|<span data-ttu-id="c6f5f-1145">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1145">Type</span></span>|<span data-ttu-id="c6f5f-1146">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1146">Attributes</span></span>|<span data-ttu-id="c6f5f-1147">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1147">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c6f5f-1148">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1148">Object</span></span>|<span data-ttu-id="c6f5f-1149">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1149">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1150">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1150">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c6f5f-1151">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1151">Object</span></span>|<span data-ttu-id="c6f5f-1152">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1152">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1153">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1153">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-1154">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1154">function</span></span>|<span data-ttu-id="c6f5f-1155">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1155">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1156">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1156">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1157">Requirements</span></span>

|<span data-ttu-id="c6f5f-1158">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1158">Requirement</span></span>|<span data-ttu-id="c6f5f-1159">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1159">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1161">1,8</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1161">1.8</span></span>|
|[<span data-ttu-id="c6f5f-1162">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1162">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1163">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1163">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1164">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1165">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1165">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6f5f-1166">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1166">Returns:</span></span>

<span data-ttu-id="c6f5f-1167">Tipo: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="c6f5f-1167">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

##### <a name="example"></a><span data-ttu-id="c6f5f-1168">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1168">Example</span></span>

<span data-ttu-id="c6f5f-1169">O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1169">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a><span data-ttu-id="c6f5f-1170">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1170">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="c6f5f-1171">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1171">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-1172">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1172">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1173">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1173">Requirements</span></span>

|<span data-ttu-id="c6f5f-1174">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1174">Requirement</span></span>|<span data-ttu-id="c6f5f-1175">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1175">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1176">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1177">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1177">1.0</span></span>|
|[<span data-ttu-id="c6f5f-1178">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1178">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1179">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1181">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1181">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6f5f-1182">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1182">Returns:</span></span>

<span data-ttu-id="c6f5f-1183">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1183">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="c6f5f-1184">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1184">Example</span></span>

<span data-ttu-id="c6f5f-1185">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1185">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a><span data-ttu-id="c6f5f-1186">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1186">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span></span>

<span data-ttu-id="c6f5f-1187">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1187">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-1188">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1188">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1189">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1189">Parameters</span></span>

|<span data-ttu-id="c6f5f-1190">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1190">Name</span></span>|<span data-ttu-id="c6f5f-1191">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1191">Type</span></span>|<span data-ttu-id="c6f5f-1192">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1192">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="c6f5f-1193">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1193">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.8)|<span data-ttu-id="c6f5f-1194">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1194">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1195">Requirements</span></span>

|<span data-ttu-id="c6f5f-1196">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1196">Requirement</span></span>|<span data-ttu-id="c6f5f-1197">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1197">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1199">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1199">1.0</span></span>|
|[<span data-ttu-id="c6f5f-1200">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1201">Restrito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1201">Restricted</span></span>|
|[<span data-ttu-id="c6f5f-1202">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1203">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1203">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6f5f-1204">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1204">Returns:</span></span>

<span data-ttu-id="c6f5f-1205">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1205">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c6f5f-1206">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1206">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="c6f5f-1207">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1207">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c6f5f-1208">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1208">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="c6f5f-1209">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1209">Value of `entityType`</span></span>|<span data-ttu-id="c6f5f-1210">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1210">Type of objects in returned array</span></span>|<span data-ttu-id="c6f5f-1211">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1211">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="c6f5f-1212">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1212">String</span></span>|<span data-ttu-id="c6f5f-1213">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1213">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="c6f5f-1214">Contato</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1214">Contact</span></span>|<span data-ttu-id="c6f5f-1215">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1215">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="c6f5f-1216">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1216">String</span></span>|<span data-ttu-id="c6f5f-1217">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1217">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="c6f5f-1218">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1218">MeetingSuggestion</span></span>|<span data-ttu-id="c6f5f-1219">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1219">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="c6f5f-1220">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1220">PhoneNumber</span></span>|<span data-ttu-id="c6f5f-1221">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1221">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="c6f5f-1222">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1222">TaskSuggestion</span></span>|<span data-ttu-id="c6f5f-1223">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1223">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="c6f5f-1224">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1224">String</span></span>|<span data-ttu-id="c6f5f-1225">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1225">**Restricted**</span></span>|

<span data-ttu-id="c6f5f-1226">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span><span class="sxs-lookup"><span data-stu-id="c6f5f-1226">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span></span>

##### <a name="example"></a><span data-ttu-id="c6f5f-1227">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1227">Example</span></span>

<span data-ttu-id="c6f5f-1228">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1228">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a><span data-ttu-id="c6f5f-1229">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1229">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span></span>

<span data-ttu-id="c6f5f-1230">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1230">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-1231">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1231">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c6f5f-1232">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1232">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1233">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1233">Parameters</span></span>

|<span data-ttu-id="c6f5f-1234">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1234">Name</span></span>|<span data-ttu-id="c6f5f-1235">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1235">Type</span></span>|<span data-ttu-id="c6f5f-1236">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1236">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c6f5f-1237">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1237">String</span></span>|<span data-ttu-id="c6f5f-1238">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1238">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1239">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1239">Requirements</span></span>

|<span data-ttu-id="c6f5f-1240">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1240">Requirement</span></span>|<span data-ttu-id="c6f5f-1241">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1241">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1242">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1243">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1243">1.0</span></span>|
|[<span data-ttu-id="c6f5f-1244">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1245">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1245">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1246">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1247">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1247">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6f5f-1248">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1248">Returns:</span></span>

<span data-ttu-id="c6f5f-p174">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c6f5f-1251">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span><span class="sxs-lookup"><span data-stu-id="c6f5f-1251">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span></span>

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="c6f5f-1252">getItemIdAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1252">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="c6f5f-1253">Obtém de forma assíncrona a ID de um item salvo.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1253">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="c6f5f-1254">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1254">Compose mode only.</span></span>

<span data-ttu-id="c6f5f-1255">Quando invocado, este método retorna a ID do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1255">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-1256">Se seu suplemento chamar `getItemIdAsync` um item no modo de redação (por exemplo, para `itemId` usar com o EWS ou a API REST), lembre-se de que, quando o Outlook estiver no modo cache, pode levar algum tempo para que o item seja sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1256">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="c6f5f-1257">Até que o item seja sincronizado, `itemId` o não é reconhecido e usado retorna um erro.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1257">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1258">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1258">Parameters</span></span>

|<span data-ttu-id="c6f5f-1259">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1259">Name</span></span>|<span data-ttu-id="c6f5f-1260">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1260">Type</span></span>|<span data-ttu-id="c6f5f-1261">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1261">Attributes</span></span>|<span data-ttu-id="c6f5f-1262">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1262">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c6f5f-1263">Object</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1263">Object</span></span>|<span data-ttu-id="c6f5f-1264">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1264">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1265">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1265">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c6f5f-1266">Object</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1266">Object</span></span>|<span data-ttu-id="c6f5f-1267">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1268">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1268">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-1269">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1269">function</span></span>||<span data-ttu-id="c6f5f-1270">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1270">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c6f5f-1271">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1271">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c6f5f-1272">Erros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1272">Errors</span></span>

|<span data-ttu-id="c6f5f-1273">Código de erro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1273">Error code</span></span>|<span data-ttu-id="c6f5f-1274">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1274">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="c6f5f-1275">A ID não pode ser recuperada até que o item seja salvo.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1275">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1276">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1276">Requirements</span></span>

|<span data-ttu-id="c6f5f-1277">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1277">Requirement</span></span>|<span data-ttu-id="c6f5f-1278">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1278">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1279">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1280">1,8</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1280">1.8</span></span>|
|[<span data-ttu-id="c6f5f-1281">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1281">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1282">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1283">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1283">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1284">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1284">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c6f5f-1285">Exemplos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1285">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="c6f5f-1286">O exemplo a seguir mostra a estrutura do `result` parâmetro que é passado para a função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1286">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="c6f5f-1287">A `value` propriedade contém a ID do item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1287">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="c6f5f-1288">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1288">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c6f5f-1289">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1289">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-1290">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1290">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c6f5f-p178">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c6f5f-1294">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1294">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c6f5f-1295">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1295">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c6f5f-p179">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1299">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1299">Requirements</span></span>

|<span data-ttu-id="c6f5f-1300">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1300">Requirement</span></span>|<span data-ttu-id="c6f5f-1301">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1301">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1302">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1303">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1303">1.0</span></span>|
|[<span data-ttu-id="c6f5f-1304">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1305">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1306">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1307">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1307">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6f5f-1308">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1308">Returns:</span></span>

<span data-ttu-id="c6f5f-p180">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="c6f5f-1311">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1311">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c6f5f-1312">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1312">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c6f5f-1313">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1313">Example</span></span>

<span data-ttu-id="c6f5f-1314">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1314">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c6f5f-1315">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1315">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c6f5f-1316">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1316">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-1317">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1317">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c6f5f-1318">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1318">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c6f5f-p181">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1321">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1321">Parameters</span></span>

|<span data-ttu-id="c6f5f-1322">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1322">Name</span></span>|<span data-ttu-id="c6f5f-1323">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1323">Type</span></span>|<span data-ttu-id="c6f5f-1324">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1324">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c6f5f-1325">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1325">String</span></span>|<span data-ttu-id="c6f5f-1326">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1326">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1327">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1327">Requirements</span></span>

|<span data-ttu-id="c6f5f-1328">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1328">Requirement</span></span>|<span data-ttu-id="c6f5f-1329">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1329">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1330">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1330">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1331">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1331">1.0</span></span>|
|[<span data-ttu-id="c6f5f-1332">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1332">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1333">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1333">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1334">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1334">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1335">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1335">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6f5f-1336">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1336">Returns:</span></span>

<span data-ttu-id="c6f5f-1337">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1337">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="c6f5f-1338">Tipo: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="c6f5f-1338">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="c6f5f-1339">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1339">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c6f5f-1340">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1340">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c6f5f-1341">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1341">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c6f5f-p182">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna uma cadeia de caracteres vazia para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p182">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1344">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1344">Parameters</span></span>

|<span data-ttu-id="c6f5f-1345">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1345">Name</span></span>|<span data-ttu-id="c6f5f-1346">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1346">Type</span></span>|<span data-ttu-id="c6f5f-1347">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1347">Attributes</span></span>|<span data-ttu-id="c6f5f-1348">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1348">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="c6f5f-1349">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1349">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c6f5f-p183">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p183">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="c6f5f-1353">Object</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1353">Object</span></span>|<span data-ttu-id="c6f5f-1354">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1354">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1355">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1355">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c6f5f-1356">Object</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1356">Object</span></span>|<span data-ttu-id="c6f5f-1357">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1357">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1358">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1358">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-1359">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1359">function</span></span>||<span data-ttu-id="c6f5f-1360">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1360">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c6f5f-1361">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1361">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c6f5f-1362">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1362">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1363">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1363">Requirements</span></span>

|<span data-ttu-id="c6f5f-1364">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1364">Requirement</span></span>|<span data-ttu-id="c6f5f-1365">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1365">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1366">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1366">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1367">1.2</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1367">1.2</span></span>|
|[<span data-ttu-id="c6f5f-1368">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1368">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1369">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1369">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1370">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1370">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1371">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1371">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6f5f-1372">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1372">Returns:</span></span>

<span data-ttu-id="c6f5f-1373">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1373">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="c6f5f-1374">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1374">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c6f5f-1375">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1375">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a><span data-ttu-id="c6f5f-1376">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1376">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="c6f5f-1377">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1377">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="c6f5f-1378">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1378">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-1379">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1379">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1380">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1380">Requirements</span></span>

|<span data-ttu-id="c6f5f-1381">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1381">Requirement</span></span>|<span data-ttu-id="c6f5f-1382">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1382">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1383">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1383">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1384">1.6</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1384">1.6</span></span>|
|[<span data-ttu-id="c6f5f-1385">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1385">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1386">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1386">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1387">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1387">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1388">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1388">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6f5f-1389">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1389">Returns:</span></span>

<span data-ttu-id="c6f5f-1390">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1390">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="c6f5f-1391">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1391">Example</span></span>

<span data-ttu-id="c6f5f-1392">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1392">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="c6f5f-1393">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1393">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="c6f5f-p186">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p186">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-1396">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1396">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c6f5f-p187">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p187">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c6f5f-1400">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1400">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c6f5f-1401">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1401">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c6f5f-p188">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p188">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1405">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1405">Requirements</span></span>

|<span data-ttu-id="c6f5f-1406">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1406">Requirement</span></span>|<span data-ttu-id="c6f5f-1407">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1407">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1408">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1409">1.6</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1409">1.6</span></span>|
|[<span data-ttu-id="c6f5f-1410">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1410">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1411">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1412">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1412">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1413">Read</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1413">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6f5f-1414">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1414">Returns:</span></span>

<span data-ttu-id="c6f5f-p189">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p189">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="c6f5f-1417">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1417">Example</span></span>

<span data-ttu-id="c6f5f-1418">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1418">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="c6f5f-1419">getSharedPropertiesAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1419">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="c6f5f-1420">Obtém as propriedades do compromisso ou da mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1420">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1421">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1421">Parameters</span></span>

|<span data-ttu-id="c6f5f-1422">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1422">Name</span></span>|<span data-ttu-id="c6f5f-1423">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1423">Type</span></span>|<span data-ttu-id="c6f5f-1424">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1424">Attributes</span></span>|<span data-ttu-id="c6f5f-1425">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1425">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c6f5f-1426">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1426">Object</span></span>|<span data-ttu-id="c6f5f-1427">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1427">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1428">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1428">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c6f5f-1429">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1429">Object</span></span>|<span data-ttu-id="c6f5f-1430">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1430">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1431">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1431">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-1432">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1432">function</span></span>||<span data-ttu-id="c6f5f-1433">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1433">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c6f5f-1434">As propriedades compartilhadas são fornecidas [`SharedProperties`](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) como um objeto `asyncResult.value` na propriedade.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1434">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c6f5f-1435">Este objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1435">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1436">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1436">Requirements</span></span>

|<span data-ttu-id="c6f5f-1437">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1437">Requirement</span></span>|<span data-ttu-id="c6f5f-1438">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1438">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1439">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1440">1,8</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1440">1.8</span></span>|
|[<span data-ttu-id="c6f5f-1441">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1441">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1442">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1443">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1443">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1444">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1444">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-1445">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1445">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c6f5f-1446">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1446">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c6f5f-1447">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1447">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c6f5f-p191">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p191">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1451">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1451">Parameters</span></span>

|<span data-ttu-id="c6f5f-1452">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1452">Name</span></span>|<span data-ttu-id="c6f5f-1453">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1453">Type</span></span>|<span data-ttu-id="c6f5f-1454">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1454">Attributes</span></span>|<span data-ttu-id="c6f5f-1455">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1455">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="c6f5f-1456">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1456">function</span></span>||<span data-ttu-id="c6f5f-1457">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1457">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c6f5f-1458">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.8) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1458">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.8) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c6f5f-1459">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1459">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="c6f5f-1460">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1460">Object</span></span>|<span data-ttu-id="c6f5f-1461">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1461">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1462">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1462">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="c6f5f-1463">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1463">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1464">Requirements</span></span>

|<span data-ttu-id="c6f5f-1465">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1465">Requirement</span></span>|<span data-ttu-id="c6f5f-1466">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1466">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1467">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1468">1.0</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1468">1.0</span></span>|
|[<span data-ttu-id="c6f5f-1469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1470">ReadItem</span></span>|
|[<span data-ttu-id="c6f5f-1471">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1472">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1472">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-1473">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1473">Example</span></span>

<span data-ttu-id="c6f5f-p194">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p194">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c6f5f-1477">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1477">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c6f5f-1478">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1478">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c6f5f-1479">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1479">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="c6f5f-1480">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1480">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="c6f5f-1481">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1481">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="c6f5f-1482">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1482">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1483">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1483">Parameters</span></span>

|<span data-ttu-id="c6f5f-1484">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1484">Name</span></span>|<span data-ttu-id="c6f5f-1485">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1485">Type</span></span>|<span data-ttu-id="c6f5f-1486">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1486">Attributes</span></span>|<span data-ttu-id="c6f5f-1487">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1487">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="c6f5f-1488">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1488">String</span></span>||<span data-ttu-id="c6f5f-1489">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1489">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="c6f5f-1490">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1490">Object</span></span>|<span data-ttu-id="c6f5f-1491">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1491">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1492">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1492">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c6f5f-1493">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1493">Object</span></span>|<span data-ttu-id="c6f5f-1494">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1494">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1495">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1495">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-1496">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1496">function</span></span>|<span data-ttu-id="c6f5f-1497">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1497">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1498">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1498">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c6f5f-1499">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1499">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c6f5f-1500">Erros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1500">Errors</span></span>

|<span data-ttu-id="c6f5f-1501">Código de erro</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1501">Error code</span></span>|<span data-ttu-id="c6f5f-1502">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1502">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="c6f5f-1503">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1503">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1504">Requirements</span></span>

|<span data-ttu-id="c6f5f-1505">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1505">Requirement</span></span>|<span data-ttu-id="c6f5f-1506">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1506">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1507">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1508">1.1</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1508">1.1</span></span>|
|[<span data-ttu-id="c6f5f-1509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1510">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1510">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6f5f-1511">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1512">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1512">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-1513">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1513">Example</span></span>

<span data-ttu-id="c6f5f-1514">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1514">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="c6f5f-1515">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1515">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="c6f5f-1516">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1516">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="c6f5f-1517">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1517">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1518">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1518">Parameters</span></span>

| <span data-ttu-id="c6f5f-1519">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1519">Name</span></span> | <span data-ttu-id="c6f5f-1520">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1520">Type</span></span> | <span data-ttu-id="c6f5f-1521">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1521">Attributes</span></span> | <span data-ttu-id="c6f5f-1522">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1522">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c6f5f-1523">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1523">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c6f5f-1524">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1524">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="c6f5f-1525">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1525">Object</span></span> | <span data-ttu-id="c6f5f-1526">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1526">&lt;optional&gt;</span></span> | <span data-ttu-id="c6f5f-1527">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1527">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c6f5f-1528">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1528">Object</span></span> | <span data-ttu-id="c6f5f-1529">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1529">&lt;optional&gt;</span></span> | <span data-ttu-id="c6f5f-1530">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1530">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c6f5f-1531">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1531">function</span></span>| <span data-ttu-id="c6f5f-1532">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1532">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1533">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1533">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1534">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1534">Requirements</span></span>

|<span data-ttu-id="c6f5f-1535">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1535">Requirement</span></span>| <span data-ttu-id="c6f5f-1536">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1536">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1537">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6f5f-1538">1.7</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1538">1.7</span></span> |
|[<span data-ttu-id="c6f5f-1539">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1539">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6f5f-1540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1540">ReadItem</span></span> |
|[<span data-ttu-id="c6f5f-1541">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1541">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6f5f-1542">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1542">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="c6f5f-1543">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1543">saveAsync([options], callback)</span></span>

<span data-ttu-id="c6f5f-1544">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1544">Asynchronously saves an item.</span></span>

<span data-ttu-id="c6f5f-1545">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1545">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="c6f5f-1546">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1546">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="c6f5f-1547">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1547">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-1548">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1548">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="c6f5f-1549">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1549">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="c6f5f-p198">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p198">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="c6f5f-1553">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1553">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="c6f5f-1554">O Outlook no Mac não dá suporte ao salvamento de reuniões.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1554">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="c6f5f-1555">O método `saveAsync` falha quando chamado a partir de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1555">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="c6f5f-1556">Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1556">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="c6f5f-1557">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1557">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1558">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1558">Parameters</span></span>

|<span data-ttu-id="c6f5f-1559">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1559">Name</span></span>|<span data-ttu-id="c6f5f-1560">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1560">Type</span></span>|<span data-ttu-id="c6f5f-1561">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1561">Attributes</span></span>|<span data-ttu-id="c6f5f-1562">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1562">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c6f5f-1563">Object</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1563">Object</span></span>|<span data-ttu-id="c6f5f-1564">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1564">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1565">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1565">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c6f5f-1566">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1566">Object</span></span>|<span data-ttu-id="c6f5f-1567">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1567">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1568">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1568">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-1569">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1569">function</span></span>||<span data-ttu-id="c6f5f-1570">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1570">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c6f5f-1571">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1571">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1572">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1572">Requirements</span></span>

|<span data-ttu-id="c6f5f-1573">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1573">Requirement</span></span>|<span data-ttu-id="c6f5f-1574">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1574">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1575">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1575">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1576">1.3</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1576">1.3</span></span>|
|[<span data-ttu-id="c6f5f-1577">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1577">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1578">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1578">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6f5f-1579">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1579">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1580">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1580">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c6f5f-1581">Exemplos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1581">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="c6f5f-p200">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p200">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c6f5f-1584">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1584">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c6f5f-1585">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1585">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c6f5f-p201">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p201">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6f5f-1589">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1589">Parameters</span></span>

|<span data-ttu-id="c6f5f-1590">Nome</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1590">Name</span></span>|<span data-ttu-id="c6f5f-1591">Tipo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1591">Type</span></span>|<span data-ttu-id="c6f5f-1592">Atributos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1592">Attributes</span></span>|<span data-ttu-id="c6f5f-1593">Descrição</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1593">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="c6f5f-1594">String</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1594">String</span></span>||<span data-ttu-id="c6f5f-p202">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-p202">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="c6f5f-1598">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1598">Object</span></span>|<span data-ttu-id="c6f5f-1599">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1599">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1600">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1600">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c6f5f-1601">Objeto</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1601">Object</span></span>|<span data-ttu-id="c6f5f-1602">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1602">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1603">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1603">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="c6f5f-1604">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1604">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="c6f5f-1605">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1605">&lt;optional&gt;</span></span>|<span data-ttu-id="c6f5f-1606">Se `text`, o estilo atual é aplicado nos clientes do Outlook na Web e do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1606">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="c6f5f-1607">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1607">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c6f5f-1608">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook na Web e o estilo padrão é aplicado nos clientes do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1608">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="c6f5f-1609">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1609">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c6f5f-1610">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1610">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="c6f5f-1611">function</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1611">function</span></span>||<span data-ttu-id="c6f5f-1612">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1612">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6f5f-1613">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1613">Requirements</span></span>

|<span data-ttu-id="c6f5f-1614">Requisito</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1614">Requirement</span></span>|<span data-ttu-id="c6f5f-1615">Valor</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1615">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6f5f-1616">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1616">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c6f5f-1617">1.2</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1617">1.2</span></span>|
|[<span data-ttu-id="c6f5f-1618">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1618">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c6f5f-1619">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1619">ReadWriteItem</span></span>|
|[<span data-ttu-id="c6f5f-1620">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1620">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c6f5f-1621">Escrever</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1621">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c6f5f-1622">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c6f5f-1622">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
