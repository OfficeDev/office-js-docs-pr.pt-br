---
title: Office. Context. Mailbox. Item-visualização do conjunto de requisitos
description: ''
ms.date: 11/06/2019
localization_priority: Normal
ms.openlocfilehash: 8a65f3b36c6c05c6885cb6925b61ee8c9520dc4a
ms.sourcegitcommit: 08c0b9ff319c391922fa43d3c2e9783cf6b53b1b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/08/2019
ms.locfileid: "38066288"
---
# <a name="item"></a><span data-ttu-id="27d2a-102">item</span><span class="sxs-lookup"><span data-stu-id="27d2a-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="27d2a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="27d2a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="27d2a-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="27d2a-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-106">Requirements</span></span>

|<span data-ttu-id="27d2a-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-107">Requirement</span></span>|<span data-ttu-id="27d2a-108">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-110">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-110">1.0</span></span>|
|[<span data-ttu-id="27d2a-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="27d2a-112">Restricted</span></span>|
|[<span data-ttu-id="27d2a-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="27d2a-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="27d2a-115">Members and methods</span></span>

| <span data-ttu-id="27d2a-116">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-116">Member</span></span> | <span data-ttu-id="27d2a-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="27d2a-118">attachments</span><span class="sxs-lookup"><span data-stu-id="27d2a-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="27d2a-119">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-119">Member</span></span> |
| [<span data-ttu-id="27d2a-120">bcc</span><span class="sxs-lookup"><span data-stu-id="27d2a-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="27d2a-121">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-121">Member</span></span> |
| [<span data-ttu-id="27d2a-122">body</span><span class="sxs-lookup"><span data-stu-id="27d2a-122">body</span></span>](#body-body) | <span data-ttu-id="27d2a-123">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-123">Member</span></span> |
| [<span data-ttu-id="27d2a-124">Categorias</span><span class="sxs-lookup"><span data-stu-id="27d2a-124">categories</span></span>](#categories-categories) | <span data-ttu-id="27d2a-125">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-125">Member</span></span> |
| [<span data-ttu-id="27d2a-126">cc</span><span class="sxs-lookup"><span data-stu-id="27d2a-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="27d2a-127">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-127">Member</span></span> |
| [<span data-ttu-id="27d2a-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="27d2a-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="27d2a-129">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-129">Member</span></span> |
| [<span data-ttu-id="27d2a-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="27d2a-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="27d2a-131">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-131">Member</span></span> |
| [<span data-ttu-id="27d2a-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="27d2a-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="27d2a-133">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-133">Member</span></span> |
| [<span data-ttu-id="27d2a-134">end</span><span class="sxs-lookup"><span data-stu-id="27d2a-134">end</span></span>](#end-datetime) | <span data-ttu-id="27d2a-135">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-135">Member</span></span> |
| [<span data-ttu-id="27d2a-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="27d2a-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="27d2a-137">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-137">Member</span></span> |
| [<span data-ttu-id="27d2a-138">from</span><span class="sxs-lookup"><span data-stu-id="27d2a-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="27d2a-139">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-139">Member</span></span> |
| [<span data-ttu-id="27d2a-140">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="27d2a-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="27d2a-141">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-141">Member</span></span> |
| [<span data-ttu-id="27d2a-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="27d2a-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="27d2a-143">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-143">Member</span></span> |
| [<span data-ttu-id="27d2a-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="27d2a-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="27d2a-145">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-145">Member</span></span> |
| [<span data-ttu-id="27d2a-146">itemId</span><span class="sxs-lookup"><span data-stu-id="27d2a-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="27d2a-147">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-147">Member</span></span> |
| [<span data-ttu-id="27d2a-148">itemType</span><span class="sxs-lookup"><span data-stu-id="27d2a-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="27d2a-149">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-149">Member</span></span> |
| [<span data-ttu-id="27d2a-150">location</span><span class="sxs-lookup"><span data-stu-id="27d2a-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="27d2a-151">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-151">Member</span></span> |
| [<span data-ttu-id="27d2a-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="27d2a-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="27d2a-153">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-153">Member</span></span> |
| [<span data-ttu-id="27d2a-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="27d2a-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="27d2a-155">Member</span><span class="sxs-lookup"><span data-stu-id="27d2a-155">Member</span></span> |
| [<span data-ttu-id="27d2a-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="27d2a-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="27d2a-157">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-157">Member</span></span> |
| [<span data-ttu-id="27d2a-158">organizer</span><span class="sxs-lookup"><span data-stu-id="27d2a-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="27d2a-159">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-159">Member</span></span> |
| [<span data-ttu-id="27d2a-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="27d2a-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="27d2a-161">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-161">Member</span></span> |
| [<span data-ttu-id="27d2a-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="27d2a-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="27d2a-163">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-163">Member</span></span> |
| [<span data-ttu-id="27d2a-164">sender</span><span class="sxs-lookup"><span data-stu-id="27d2a-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="27d2a-165">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-165">Member</span></span> |
| [<span data-ttu-id="27d2a-166">seriesid</span><span class="sxs-lookup"><span data-stu-id="27d2a-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="27d2a-167">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-167">Member</span></span> |
| [<span data-ttu-id="27d2a-168">start</span><span class="sxs-lookup"><span data-stu-id="27d2a-168">start</span></span>](#start-datetime) | <span data-ttu-id="27d2a-169">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-169">Member</span></span> |
| [<span data-ttu-id="27d2a-170">subject</span><span class="sxs-lookup"><span data-stu-id="27d2a-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="27d2a-171">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-171">Member</span></span> |
| [<span data-ttu-id="27d2a-172">to</span><span class="sxs-lookup"><span data-stu-id="27d2a-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="27d2a-173">Membro</span><span class="sxs-lookup"><span data-stu-id="27d2a-173">Member</span></span> |
| [<span data-ttu-id="27d2a-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="27d2a-175">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-175">Method</span></span> |
| [<span data-ttu-id="27d2a-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="27d2a-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="27d2a-177">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-177">Method</span></span> |
| [<span data-ttu-id="27d2a-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="27d2a-179">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-179">Method</span></span> |
| [<span data-ttu-id="27d2a-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="27d2a-181">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-181">Method</span></span> |
| [<span data-ttu-id="27d2a-182">close</span><span class="sxs-lookup"><span data-stu-id="27d2a-182">close</span></span>](#close) | <span data-ttu-id="27d2a-183">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-183">Method</span></span> |
| [<span data-ttu-id="27d2a-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="27d2a-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="27d2a-185">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-185">Method</span></span> |
| [<span data-ttu-id="27d2a-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="27d2a-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="27d2a-187">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-187">Method</span></span> |
| [<span data-ttu-id="27d2a-188">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-188">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="27d2a-189">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-189">Method</span></span> |
| [<span data-ttu-id="27d2a-190">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-190">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="27d2a-191">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-191">Method</span></span> |
| [<span data-ttu-id="27d2a-192">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-192">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="27d2a-193">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-193">Method</span></span> |
| [<span data-ttu-id="27d2a-194">getEntities</span><span class="sxs-lookup"><span data-stu-id="27d2a-194">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="27d2a-195">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-195">Method</span></span> |
| [<span data-ttu-id="27d2a-196">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="27d2a-196">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="27d2a-197">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-197">Method</span></span> |
| [<span data-ttu-id="27d2a-198">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="27d2a-198">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="27d2a-199">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-199">Method</span></span> |
| [<span data-ttu-id="27d2a-200">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-200">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="27d2a-201">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-201">Method</span></span> |
| [<span data-ttu-id="27d2a-202">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-202">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="27d2a-203">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-203">Method</span></span> |
| [<span data-ttu-id="27d2a-204">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="27d2a-204">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="27d2a-205">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-205">Method</span></span> |
| [<span data-ttu-id="27d2a-206">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="27d2a-206">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="27d2a-207">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-207">Method</span></span> |
| [<span data-ttu-id="27d2a-208">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-208">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="27d2a-209">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-209">Method</span></span> |
| [<span data-ttu-id="27d2a-210">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="27d2a-210">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="27d2a-211">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-211">Method</span></span> |
| [<span data-ttu-id="27d2a-212">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="27d2a-212">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="27d2a-213">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-213">Method</span></span> |
| [<span data-ttu-id="27d2a-214">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-214">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="27d2a-215">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-215">Method</span></span> |
| [<span data-ttu-id="27d2a-216">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-216">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="27d2a-217">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-217">Method</span></span> |
| [<span data-ttu-id="27d2a-218">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-218">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="27d2a-219">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-219">Method</span></span> |
| [<span data-ttu-id="27d2a-220">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-220">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="27d2a-221">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-221">Method</span></span> |
| [<span data-ttu-id="27d2a-222">saveAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-222">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="27d2a-223">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-223">Method</span></span> |
| [<span data-ttu-id="27d2a-224">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="27d2a-224">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="27d2a-225">Método</span><span class="sxs-lookup"><span data-stu-id="27d2a-225">Method</span></span> |

### <a name="example"></a><span data-ttu-id="27d2a-226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-226">Example</span></span>

<span data-ttu-id="27d2a-227">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="27d2a-227">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="27d2a-228">Members</span><span class="sxs-lookup"><span data-stu-id="27d2a-228">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="27d2a-229">anexos: Matriz.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="27d2a-229">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="27d2a-230">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="27d2a-230">Gets the item's attachments as an array.</span></span> <span data-ttu-id="27d2a-231">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="27d2a-231">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-232">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="27d2a-232">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="27d2a-233">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="27d2a-233">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-234">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-234">Type</span></span>

*   <span data-ttu-id="27d2a-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="27d2a-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-236">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-236">Requirements</span></span>

|<span data-ttu-id="27d2a-237">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-237">Requirement</span></span>|<span data-ttu-id="27d2a-238">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-239">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-240">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-240">1.0</span></span>|
|[<span data-ttu-id="27d2a-241">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-242">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-243">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-244">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-244">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-245">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-245">Example</span></span>

<span data-ttu-id="27d2a-246">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="27d2a-246">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="27d2a-247">cco :[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="27d2a-247">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="27d2a-248">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="27d2a-248">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="27d2a-249">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="27d2a-249">Compose mode only.</span></span>

<span data-ttu-id="27d2a-250">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="27d2a-251">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="27d2a-251">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="27d2a-252">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-252">Get 500 members maximum.</span></span>
- <span data-ttu-id="27d2a-253">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="27d2a-253">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-254">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-254">Type</span></span>

*   [<span data-ttu-id="27d2a-255">Destinatários</span><span class="sxs-lookup"><span data-stu-id="27d2a-255">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="27d2a-256">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-256">Requirements</span></span>

|<span data-ttu-id="27d2a-257">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-257">Requirement</span></span>|<span data-ttu-id="27d2a-258">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-259">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-260">1.1</span><span class="sxs-lookup"><span data-stu-id="27d2a-260">1.1</span></span>|
|[<span data-ttu-id="27d2a-261">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-262">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-264">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-264">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-265">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-265">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="27d2a-266">corpo: [Corpo](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="27d2a-266">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="27d2a-267">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-267">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-268">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-268">Type</span></span>

*   [<span data-ttu-id="27d2a-269">Body</span><span class="sxs-lookup"><span data-stu-id="27d2a-269">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="27d2a-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-270">Requirements</span></span>

|<span data-ttu-id="27d2a-271">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-271">Requirement</span></span>|<span data-ttu-id="27d2a-272">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-273">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-274">1.1</span><span class="sxs-lookup"><span data-stu-id="27d2a-274">1.1</span></span>|
|[<span data-ttu-id="27d2a-275">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-275">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-276">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-277">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-277">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-278">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-279">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-279">Example</span></span>

<span data-ttu-id="27d2a-280">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="27d2a-280">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="27d2a-281">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-281">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="27d2a-282">Categorias: [categorias](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="27d2a-282">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="27d2a-283">Obtém um objeto que fornece métodos para gerenciar as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-283">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-284">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="27d2a-284">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-285">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-285">Type</span></span>

*   [<span data-ttu-id="27d2a-286">Categories</span><span class="sxs-lookup"><span data-stu-id="27d2a-286">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="27d2a-287">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-287">Requirements</span></span>

|<span data-ttu-id="27d2a-288">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-288">Requirement</span></span>|<span data-ttu-id="27d2a-289">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-290">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-291">1,8</span><span class="sxs-lookup"><span data-stu-id="27d2a-291">1.8</span></span>|
|[<span data-ttu-id="27d2a-292">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-293">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-293">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-294">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-295">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-295">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-296">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-296">Example</span></span>

<span data-ttu-id="27d2a-297">Este exemplo obtém as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-297">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="27d2a-298">cc : Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="27d2a-298">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="27d2a-299">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="27d2a-299">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="27d2a-300">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="27d2a-300">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="27d2a-301">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="27d2a-301">Read mode</span></span>

<span data-ttu-id="27d2a-302">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="27d2a-302">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="27d2a-303">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-303">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="27d2a-304">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-304">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="27d2a-305">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="27d2a-305">Compose mode</span></span>

<span data-ttu-id="27d2a-306">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="27d2a-306">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="27d2a-307">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-307">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="27d2a-308">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="27d2a-308">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="27d2a-309">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-309">Get 500 members maximum.</span></span>
- <span data-ttu-id="27d2a-310">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="27d2a-310">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="27d2a-311">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-311">Type</span></span>

*   <span data-ttu-id="27d2a-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="27d2a-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-313">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-313">Requirements</span></span>

|<span data-ttu-id="27d2a-314">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-314">Requirement</span></span>|<span data-ttu-id="27d2a-315">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-316">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-317">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-317">1.0</span></span>|
|[<span data-ttu-id="27d2a-318">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-319">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-320">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-321">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-321">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="27d2a-322">(anulável) conversationId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-322">(nullable) conversationId: String</span></span>

<span data-ttu-id="27d2a-323">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="27d2a-323">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="27d2a-p109">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="27d2a-p110">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-328">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-328">Type</span></span>

*   <span data-ttu-id="27d2a-329">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-329">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-330">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-330">Requirements</span></span>

|<span data-ttu-id="27d2a-331">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-331">Requirement</span></span>|<span data-ttu-id="27d2a-332">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-333">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-334">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-334">1.0</span></span>|
|[<span data-ttu-id="27d2a-335">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-336">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-337">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-338">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-338">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-339">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-339">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="27d2a-340">dateTimeCreated: Data</span><span class="sxs-lookup"><span data-stu-id="27d2a-340">dateTimeCreated: Date</span></span>

<span data-ttu-id="27d2a-p111">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-343">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-343">Type</span></span>

*   <span data-ttu-id="27d2a-344">Data</span><span class="sxs-lookup"><span data-stu-id="27d2a-344">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-345">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-345">Requirements</span></span>

|<span data-ttu-id="27d2a-346">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-346">Requirement</span></span>|<span data-ttu-id="27d2a-347">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-348">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-349">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-349">1.0</span></span>|
|[<span data-ttu-id="27d2a-350">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-350">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-351">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-352">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-352">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-353">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-353">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-354">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-354">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="27d2a-355">dateTimeModified: Data</span><span class="sxs-lookup"><span data-stu-id="27d2a-355">dateTimeModified: Date</span></span>

<span data-ttu-id="27d2a-p112">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-358">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="27d2a-358">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-359">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-359">Type</span></span>

*   <span data-ttu-id="27d2a-360">Data</span><span class="sxs-lookup"><span data-stu-id="27d2a-360">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-361">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-361">Requirements</span></span>

|<span data-ttu-id="27d2a-362">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-362">Requirement</span></span>|<span data-ttu-id="27d2a-363">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-363">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-364">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-364">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-365">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-365">1.0</span></span>|
|[<span data-ttu-id="27d2a-366">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-366">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-367">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-368">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-368">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-369">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-369">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-370">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-370">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="27d2a-371">fim: Data|[Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="27d2a-371">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="27d2a-372">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="27d2a-372">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="27d2a-p113">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="27d2a-375">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="27d2a-375">Read mode</span></span>

<span data-ttu-id="27d2a-376">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-376">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="27d2a-377">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="27d2a-377">Compose mode</span></span>

<span data-ttu-id="27d2a-378">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-378">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="27d2a-379">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="27d2a-379">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="27d2a-380">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-380">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="27d2a-381">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-381">Type</span></span>

*   <span data-ttu-id="27d2a-382">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="27d2a-382">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-383">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-383">Requirements</span></span>

|<span data-ttu-id="27d2a-384">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-384">Requirement</span></span>|<span data-ttu-id="27d2a-385">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-386">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-387">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-387">1.0</span></span>|
|[<span data-ttu-id="27d2a-388">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-389">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-390">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-391">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-391">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="27d2a-392">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="27d2a-392">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="27d2a-393">Obtém ou define os locais de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-393">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="27d2a-394">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="27d2a-394">Read mode</span></span>

<span data-ttu-id="27d2a-395">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que permite que você obtenha o conjunto de locais (cada um representado por um objeto [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associado ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-395">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="27d2a-396">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="27d2a-396">Compose mode</span></span>

<span data-ttu-id="27d2a-397">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que fornece métodos para obter, remover ou adicionar locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-397">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-398">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-398">Type</span></span>

*   [<span data-ttu-id="27d2a-399">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="27d2a-399">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="27d2a-400">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-400">Requirements</span></span>

|<span data-ttu-id="27d2a-401">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-401">Requirement</span></span>|<span data-ttu-id="27d2a-402">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-403">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-403">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-404">1,8</span><span class="sxs-lookup"><span data-stu-id="27d2a-404">1.8</span></span>|
|[<span data-ttu-id="27d2a-405">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-405">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-406">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-407">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-407">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-408">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-408">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-409">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-409">Example</span></span>

<span data-ttu-id="27d2a-410">O exemplo a seguir obtém os locais atuais associados ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-410">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="27d2a-411">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="27d2a-411">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="27d2a-412">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="27d2a-412">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="27d2a-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-415">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-415">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="27d2a-416">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="27d2a-416">Read mode</span></span>

<span data-ttu-id="27d2a-417">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="27d2a-417">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="27d2a-418">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="27d2a-418">Compose mode</span></span>

<span data-ttu-id="27d2a-419">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="27d2a-419">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="27d2a-420">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-420">Type</span></span>

*   <span data-ttu-id="27d2a-421">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="27d2a-421">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-422">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-422">Requirements</span></span>

|<span data-ttu-id="27d2a-423">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-423">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="27d2a-424">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-425">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-425">1.0</span></span>|<span data-ttu-id="27d2a-426">1.7</span><span class="sxs-lookup"><span data-stu-id="27d2a-426">1.7</span></span>|
|[<span data-ttu-id="27d2a-427">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-427">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-428">ReadItem</span></span>|<span data-ttu-id="27d2a-429">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-429">ReadWriteItem</span></span>|
|[<span data-ttu-id="27d2a-430">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-431">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-431">Read</span></span>|<span data-ttu-id="27d2a-432">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-432">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="27d2a-433">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="27d2a-433">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="27d2a-434">Obtém ou define cabeçalhos de Internet personalizados em uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="27d2a-434">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="27d2a-435">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="27d2a-435">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-436">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-436">Type</span></span>

*   [<span data-ttu-id="27d2a-437">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="27d2a-437">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="27d2a-438">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-438">Requirements</span></span>

|<span data-ttu-id="27d2a-439">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-439">Requirement</span></span>|<span data-ttu-id="27d2a-440">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-440">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-441">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-442">1,8</span><span class="sxs-lookup"><span data-stu-id="27d2a-442">1.8</span></span>|
|[<span data-ttu-id="27d2a-443">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-444">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-445">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-446">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-446">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-447">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-447">Example</span></span>

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

#### <a name="internetmessageid-string"></a><span data-ttu-id="27d2a-448">internetMessageId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-448">internetMessageId: String</span></span>

<span data-ttu-id="27d2a-p116">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-451">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-451">Type</span></span>

*   <span data-ttu-id="27d2a-452">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-453">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-453">Requirements</span></span>

|<span data-ttu-id="27d2a-454">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-454">Requirement</span></span>|<span data-ttu-id="27d2a-455">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-456">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-457">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-457">1.0</span></span>|
|[<span data-ttu-id="27d2a-458">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-458">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-459">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-460">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-460">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-461">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-462">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-462">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="27d2a-463">itemClass: Cadeia de caracteres </span><span class="sxs-lookup"><span data-stu-id="27d2a-463">itemClass: String</span></span>

<span data-ttu-id="27d2a-p117">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="27d2a-p118">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="27d2a-468">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-468">Type</span></span>|<span data-ttu-id="27d2a-469">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-469">Description</span></span>|<span data-ttu-id="27d2a-470">classe de item</span><span class="sxs-lookup"><span data-stu-id="27d2a-470">item class</span></span>|
|---|---|---|
|<span data-ttu-id="27d2a-471">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="27d2a-471">Appointment items</span></span>|<span data-ttu-id="27d2a-472">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-472">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="27d2a-473">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="27d2a-473">Message items</span></span>|<span data-ttu-id="27d2a-474">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="27d2a-474">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="27d2a-475">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-475">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-476">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-476">Type</span></span>

*   <span data-ttu-id="27d2a-477">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-477">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-478">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-478">Requirements</span></span>

|<span data-ttu-id="27d2a-479">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-479">Requirement</span></span>|<span data-ttu-id="27d2a-480">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-481">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-482">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-482">1.0</span></span>|
|[<span data-ttu-id="27d2a-483">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-484">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-485">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-486">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-486">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-487">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-487">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="27d2a-488">(anulável) itemId: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-488">(nullable) itemId: String</span></span>

<span data-ttu-id="27d2a-p119">Obtém o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-491">O identificador retornado pela propriedade `itemId` é o mesmo que o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="27d2a-491">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="27d2a-492">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="27d2a-492">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="27d2a-493">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="27d2a-493">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="27d2a-494">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="27d2a-494">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="27d2a-p121">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-497">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-497">Type</span></span>

*   <span data-ttu-id="27d2a-498">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-498">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-499">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-499">Requirements</span></span>

|<span data-ttu-id="27d2a-500">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-500">Requirement</span></span>|<span data-ttu-id="27d2a-501">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-502">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-503">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-503">1.0</span></span>|
|[<span data-ttu-id="27d2a-504">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-504">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-505">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-506">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-506">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-507">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-508">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-508">Example</span></span>

<span data-ttu-id="27d2a-p122">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="27d2a-511">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="27d2a-511">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="27d2a-512">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="27d2a-512">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="27d2a-513">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-513">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-514">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-514">Type</span></span>

*   [<span data-ttu-id="27d2a-515">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="27d2a-515">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="27d2a-516">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-516">Requirements</span></span>

|<span data-ttu-id="27d2a-517">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-517">Requirement</span></span>|<span data-ttu-id="27d2a-518">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-519">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-520">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-520">1.0</span></span>|
|[<span data-ttu-id="27d2a-521">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-521">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-522">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-523">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-523">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-524">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-524">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-525">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-525">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="27d2a-526">Local: Cadeia de caracteres[Local](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="27d2a-526">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="27d2a-527">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-527">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="27d2a-528">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="27d2a-528">Read mode</span></span>

<span data-ttu-id="27d2a-529">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-529">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="27d2a-530">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="27d2a-530">Compose mode</span></span>

<span data-ttu-id="27d2a-531">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-531">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="27d2a-532">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-532">Type</span></span>

*   <span data-ttu-id="27d2a-533">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="27d2a-533">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-534">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-534">Requirements</span></span>

|<span data-ttu-id="27d2a-535">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-535">Requirement</span></span>|<span data-ttu-id="27d2a-536">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-537">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-538">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-538">1.0</span></span>|
|[<span data-ttu-id="27d2a-539">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-539">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-540">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-541">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-541">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-542">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-542">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="27d2a-543">normalizedSubject: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-543">normalizedSubject: String</span></span>

<span data-ttu-id="27d2a-p123">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="27d2a-p124">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="27d2a-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-548">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-548">Type</span></span>

*   <span data-ttu-id="27d2a-549">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-549">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-550">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-550">Requirements</span></span>

|<span data-ttu-id="27d2a-551">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-551">Requirement</span></span>|<span data-ttu-id="27d2a-552">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-552">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-553">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-553">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-554">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-554">1.0</span></span>|
|[<span data-ttu-id="27d2a-555">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-555">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-556">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-556">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-557">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-557">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-558">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-558">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-559">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-559">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="27d2a-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="27d2a-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="27d2a-561">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-561">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-562">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-562">Type</span></span>

*   [<span data-ttu-id="27d2a-563">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="27d2a-563">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="27d2a-564">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-564">Requirements</span></span>

|<span data-ttu-id="27d2a-565">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-565">Requirement</span></span>|<span data-ttu-id="27d2a-566">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-567">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-568">1.3</span><span class="sxs-lookup"><span data-stu-id="27d2a-568">1.3</span></span>|
|[<span data-ttu-id="27d2a-569">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-570">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-571">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-572">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-572">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-573">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-573">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="27d2a-574">optionalAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="27d2a-574">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="27d2a-575">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="27d2a-575">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="27d2a-576">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="27d2a-576">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="27d2a-577">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="27d2a-577">Read mode</span></span>

<span data-ttu-id="27d2a-578">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="27d2a-578">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="27d2a-579">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-579">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="27d2a-580">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-580">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="27d2a-581">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="27d2a-581">Compose mode</span></span>

<span data-ttu-id="27d2a-582">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="27d2a-582">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="27d2a-583">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-583">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="27d2a-584">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="27d2a-584">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="27d2a-585">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-585">Get 500 members maximum.</span></span>
- <span data-ttu-id="27d2a-586">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="27d2a-586">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="27d2a-587">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-587">Type</span></span>

*   <span data-ttu-id="27d2a-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="27d2a-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-589">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-589">Requirements</span></span>

|<span data-ttu-id="27d2a-590">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-590">Requirement</span></span>|<span data-ttu-id="27d2a-591">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-592">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-593">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-593">1.0</span></span>|
|[<span data-ttu-id="27d2a-594">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-595">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-595">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-596">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-597">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-597">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="27d2a-598">organizador: [](/javascript/api/outlook/office.emailaddressdetails)|[organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="27d2a-598">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="27d2a-599">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-599">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="27d2a-600">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="27d2a-600">Read mode</span></span>

<span data-ttu-id="27d2a-601">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="27d2a-601">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="27d2a-602">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="27d2a-602">Compose mode</span></span>

<span data-ttu-id="27d2a-603">A `organizer` propriedade retorna um objeto [organizador](/javascript/api/outlook/office.organizer) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="27d2a-603">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="27d2a-604">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-604">Type</span></span>

*   <span data-ttu-id="27d2a-605">[](/javascript/api/outlook/office.emailaddressdetails) | [Organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="27d2a-605">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-606">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-606">Requirements</span></span>

|<span data-ttu-id="27d2a-607">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-607">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="27d2a-608">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-609">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-609">1.0</span></span>|<span data-ttu-id="27d2a-610">1.7</span><span class="sxs-lookup"><span data-stu-id="27d2a-610">1.7</span></span>|
|[<span data-ttu-id="27d2a-611">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-611">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-612">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-612">ReadItem</span></span>|<span data-ttu-id="27d2a-613">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-613">ReadWriteItem</span></span>|
|[<span data-ttu-id="27d2a-614">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-615">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-615">Read</span></span>|<span data-ttu-id="27d2a-616">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-616">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="27d2a-617">(anulável) recorrência: [recorrência](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="27d2a-617">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="27d2a-618">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-618">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="27d2a-619">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="27d2a-619">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="27d2a-620">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-620">Read and compose modes for appointment items.</span></span> <span data-ttu-id="27d2a-621">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="27d2a-621">Read mode for meeting request items.</span></span>

<span data-ttu-id="27d2a-622">A `recurrence` propriedade retorna um objeto de [recorrência](/javascript/api/outlook/office.recurrence) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="27d2a-622">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="27d2a-623">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="27d2a-623">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="27d2a-624">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="27d2a-624">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="27d2a-625">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="27d2a-625">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="27d2a-626">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="27d2a-626">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="27d2a-627">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="27d2a-627">Read mode</span></span>

<span data-ttu-id="27d2a-628">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-628">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="27d2a-629">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="27d2a-629">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="27d2a-630">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="27d2a-630">Compose mode</span></span>

<span data-ttu-id="27d2a-631">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-631">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="27d2a-632">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="27d2a-632">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="27d2a-633">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-633">Type</span></span>

* [<span data-ttu-id="27d2a-634">Recorrência</span><span class="sxs-lookup"><span data-stu-id="27d2a-634">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="27d2a-635">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-635">Requirement</span></span>|<span data-ttu-id="27d2a-636">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-637">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-638">1.7</span><span class="sxs-lookup"><span data-stu-id="27d2a-638">1.7</span></span>|
|[<span data-ttu-id="27d2a-639">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-640">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-641">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-642">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-642">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="27d2a-643">requiredAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="27d2a-643">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="27d2a-644">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="27d2a-644">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="27d2a-645">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="27d2a-645">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="27d2a-646">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="27d2a-646">Read mode</span></span>

<span data-ttu-id="27d2a-647">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="27d2a-647">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="27d2a-648">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-648">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="27d2a-649">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-649">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="27d2a-650">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="27d2a-650">Compose mode</span></span>

<span data-ttu-id="27d2a-651">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="27d2a-651">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="27d2a-652">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-652">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="27d2a-653">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="27d2a-653">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="27d2a-654">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-654">Get 500 members maximum.</span></span>
- <span data-ttu-id="27d2a-655">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="27d2a-655">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="27d2a-656">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-656">Type</span></span>

*   <span data-ttu-id="27d2a-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="27d2a-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-658">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-658">Requirements</span></span>

|<span data-ttu-id="27d2a-659">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-659">Requirement</span></span>|<span data-ttu-id="27d2a-660">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-661">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-662">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-662">1.0</span></span>|
|[<span data-ttu-id="27d2a-663">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-664">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-664">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-665">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-666">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-666">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="27d2a-667">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="27d2a-667">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="27d2a-p135">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="27d2a-p136">As propriedades [`from`](#from-emailaddressdetailsfrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-672">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-672">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-673">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-673">Type</span></span>

*   [<span data-ttu-id="27d2a-674">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="27d2a-674">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="27d2a-675">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-675">Requirements</span></span>

|<span data-ttu-id="27d2a-676">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-676">Requirement</span></span>|<span data-ttu-id="27d2a-677">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-677">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-678">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-678">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-679">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-679">1.0</span></span>|
|[<span data-ttu-id="27d2a-680">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-680">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-681">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-681">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-682">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-682">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-683">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-683">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-684">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-684">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="27d2a-685">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="27d2a-685">(nullable) seriesId: String</span></span>

<span data-ttu-id="27d2a-686">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="27d2a-686">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="27d2a-687">No Outlook na Web e clientes de desktop, o `seriesId` retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="27d2a-687">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="27d2a-688">No entanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="27d2a-688">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-689">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="27d2a-689">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="27d2a-690">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="27d2a-690">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="27d2a-691">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="27d2a-691">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="27d2a-692">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="27d2a-692">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="27d2a-693">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="27d2a-693">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="27d2a-694">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-694">Type</span></span>

* <span data-ttu-id="27d2a-695">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-695">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-696">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-696">Requirements</span></span>

|<span data-ttu-id="27d2a-697">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-697">Requirement</span></span>|<span data-ttu-id="27d2a-698">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-699">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-700">1.7</span><span class="sxs-lookup"><span data-stu-id="27d2a-700">1.7</span></span>|
|[<span data-ttu-id="27d2a-701">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-701">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-702">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-703">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-703">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-704">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-704">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-705">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-705">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="27d2a-706">início: Data|[Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="27d2a-706">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="27d2a-707">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="27d2a-707">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="27d2a-p139">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="27d2a-710">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="27d2a-710">Read mode</span></span>

<span data-ttu-id="27d2a-711">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-711">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="27d2a-712">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="27d2a-712">Compose mode</span></span>

<span data-ttu-id="27d2a-713">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-713">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="27d2a-714">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="27d2a-714">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="27d2a-715">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-715">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="27d2a-716">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-716">Type</span></span>

*   <span data-ttu-id="27d2a-717">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="27d2a-717">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-718">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-718">Requirements</span></span>

|<span data-ttu-id="27d2a-719">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-719">Requirement</span></span>|<span data-ttu-id="27d2a-720">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-720">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-721">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-721">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-722">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-722">1.0</span></span>|
|[<span data-ttu-id="27d2a-723">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-723">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-724">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-724">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-725">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-725">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-726">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-726">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="27d2a-727">Assunto: Cadeia de caracteres|[Assunto](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="27d2a-727">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="27d2a-728">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-728">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="27d2a-729">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="27d2a-729">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="27d2a-730">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="27d2a-730">Read mode</span></span>

<span data-ttu-id="27d2a-p140">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="27d2a-733">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="27d2a-733">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="27d2a-734">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="27d2a-734">Compose mode</span></span>
<span data-ttu-id="27d2a-735">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="27d2a-735">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="27d2a-736">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-736">Type</span></span>

*   <span data-ttu-id="27d2a-737">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="27d2a-737">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-738">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-738">Requirements</span></span>

|<span data-ttu-id="27d2a-739">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-739">Requirement</span></span>|<span data-ttu-id="27d2a-740">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-741">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-741">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-742">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-742">1.0</span></span>|
|[<span data-ttu-id="27d2a-743">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-743">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-744">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-744">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-745">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-745">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-746">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-746">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="27d2a-747">para: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="27d2a-747">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="27d2a-748">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="27d2a-748">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="27d2a-749">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="27d2a-749">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="27d2a-750">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="27d2a-750">Read mode</span></span>

<span data-ttu-id="27d2a-751">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="27d2a-751">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="27d2a-752">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-752">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="27d2a-753">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-753">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="27d2a-754">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="27d2a-754">Compose mode</span></span>

<span data-ttu-id="27d2a-755">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="27d2a-755">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="27d2a-756">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-756">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="27d2a-757">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="27d2a-757">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="27d2a-758">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="27d2a-758">Get 500 members maximum.</span></span>
- <span data-ttu-id="27d2a-759">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="27d2a-759">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="27d2a-760">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-760">Type</span></span>

*   <span data-ttu-id="27d2a-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="27d2a-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-762">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-762">Requirements</span></span>

|<span data-ttu-id="27d2a-763">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-763">Requirement</span></span>|<span data-ttu-id="27d2a-764">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-764">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-765">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-765">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-766">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-766">1.0</span></span>|
|[<span data-ttu-id="27d2a-767">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-767">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-768">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-768">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-769">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-769">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-770">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-770">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="27d2a-771">Métodos</span><span class="sxs-lookup"><span data-stu-id="27d2a-771">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="27d2a-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="27d2a-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="27d2a-773">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-773">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="27d2a-774">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="27d2a-774">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="27d2a-775">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="27d2a-775">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-776">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-776">Parameters</span></span>
|<span data-ttu-id="27d2a-777">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-777">Name</span></span>|<span data-ttu-id="27d2a-778">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-778">Type</span></span>|<span data-ttu-id="27d2a-779">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-779">Attributes</span></span>|<span data-ttu-id="27d2a-780">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-780">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="27d2a-781">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-781">String</span></span>||<span data-ttu-id="27d2a-p144">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="27d2a-784">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-784">String</span></span>||<span data-ttu-id="27d2a-p145">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="27d2a-787">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-787">Object</span></span>|<span data-ttu-id="27d2a-788">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-788">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-789">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-789">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="27d2a-790">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-790">Object</span></span>|<span data-ttu-id="27d2a-791">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-791">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-792">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-792">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="27d2a-793">Booliano</span><span class="sxs-lookup"><span data-stu-id="27d2a-793">Boolean</span></span>|<span data-ttu-id="27d2a-794">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-794">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-795">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="27d2a-795">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="27d2a-796">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-796">function</span></span>|<span data-ttu-id="27d2a-797">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-797">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-798">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-798">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="27d2a-799">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-799">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="27d2a-800">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="27d2a-800">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="27d2a-801">Erros</span><span class="sxs-lookup"><span data-stu-id="27d2a-801">Errors</span></span>

|<span data-ttu-id="27d2a-802">Código de erro</span><span class="sxs-lookup"><span data-stu-id="27d2a-802">Error code</span></span>|<span data-ttu-id="27d2a-803">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-803">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="27d2a-804">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="27d2a-804">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="27d2a-805">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="27d2a-805">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="27d2a-806">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="27d2a-806">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-807">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-807">Requirements</span></span>

|<span data-ttu-id="27d2a-808">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-808">Requirement</span></span>|<span data-ttu-id="27d2a-809">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-809">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-810">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-810">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-811">1.1</span><span class="sxs-lookup"><span data-stu-id="27d2a-811">1.1</span></span>|
|[<span data-ttu-id="27d2a-812">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-812">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-813">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-813">ReadWriteItem</span></span>|
|[<span data-ttu-id="27d2a-814">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-814">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-815">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-815">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="27d2a-816">Exemplos</span><span class="sxs-lookup"><span data-stu-id="27d2a-816">Examples</span></span>

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

<span data-ttu-id="27d2a-817">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="27d2a-817">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="27d2a-818">addFileAttachmentFromBase64Async (base64file, AttachmentName, [Options], [callback])</span><span class="sxs-lookup"><span data-stu-id="27d2a-818">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="27d2a-819">Adiciona um arquivo da codificação Base64 a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-819">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="27d2a-820">O `addFileAttachmentFromBase64Async` método carrega o arquivo da codificação Base64 e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="27d2a-820">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="27d2a-821">Esse método retorna o identificador de anexo no objeto AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="27d2a-821">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="27d2a-822">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="27d2a-822">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-823">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-823">Parameters</span></span>

|<span data-ttu-id="27d2a-824">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-824">Name</span></span>|<span data-ttu-id="27d2a-825">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-825">Type</span></span>|<span data-ttu-id="27d2a-826">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-826">Attributes</span></span>|<span data-ttu-id="27d2a-827">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-827">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="27d2a-828">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-828">String</span></span>||<span data-ttu-id="27d2a-829">O conteúdo codificado em Base64 de uma imagem ou arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="27d2a-829">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="27d2a-830">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-830">String</span></span>||<span data-ttu-id="27d2a-p147">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="27d2a-833">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-833">Object</span></span>|<span data-ttu-id="27d2a-834">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-834">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-835">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-835">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="27d2a-836">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-836">Object</span></span>|<span data-ttu-id="27d2a-837">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-837">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-838">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-838">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="27d2a-839">Booliano</span><span class="sxs-lookup"><span data-stu-id="27d2a-839">Boolean</span></span>|<span data-ttu-id="27d2a-840">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-840">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-841">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="27d2a-841">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="27d2a-842">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-842">function</span></span>|<span data-ttu-id="27d2a-843">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-843">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-844">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-844">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="27d2a-845">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-845">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="27d2a-846">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="27d2a-846">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="27d2a-847">Erros</span><span class="sxs-lookup"><span data-stu-id="27d2a-847">Errors</span></span>

|<span data-ttu-id="27d2a-848">Código de erro</span><span class="sxs-lookup"><span data-stu-id="27d2a-848">Error code</span></span>|<span data-ttu-id="27d2a-849">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-849">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="27d2a-850">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="27d2a-850">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="27d2a-851">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="27d2a-851">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="27d2a-852">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="27d2a-852">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-853">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-853">Requirements</span></span>

|<span data-ttu-id="27d2a-854">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-854">Requirement</span></span>|<span data-ttu-id="27d2a-855">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-856">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-857">1,8</span><span class="sxs-lookup"><span data-stu-id="27d2a-857">1.8</span></span>|
|[<span data-ttu-id="27d2a-858">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-858">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-859">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-859">ReadWriteItem</span></span>|
|[<span data-ttu-id="27d2a-860">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-860">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-861">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-861">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="27d2a-862">Exemplos</span><span class="sxs-lookup"><span data-stu-id="27d2a-862">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="27d2a-863">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="27d2a-863">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="27d2a-864">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="27d2a-864">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="27d2a-865">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="27d2a-865">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-866">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-866">Parameters</span></span>

| <span data-ttu-id="27d2a-867">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-867">Name</span></span> | <span data-ttu-id="27d2a-868">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-868">Type</span></span> | <span data-ttu-id="27d2a-869">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-869">Attributes</span></span> | <span data-ttu-id="27d2a-870">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-870">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="27d2a-871">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="27d2a-871">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="27d2a-872">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="27d2a-872">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="27d2a-873">Função</span><span class="sxs-lookup"><span data-stu-id="27d2a-873">Function</span></span> || <span data-ttu-id="27d2a-p148">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="27d2a-877">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-877">Object</span></span> | <span data-ttu-id="27d2a-878">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-878">&lt;optional&gt;</span></span> | <span data-ttu-id="27d2a-879">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-879">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="27d2a-880">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-880">Object</span></span> | <span data-ttu-id="27d2a-881">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-881">&lt;optional&gt;</span></span> | <span data-ttu-id="27d2a-882">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-882">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="27d2a-883">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-883">function</span></span>| <span data-ttu-id="27d2a-884">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-884">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-885">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-885">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-886">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-886">Requirements</span></span>

|<span data-ttu-id="27d2a-887">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-887">Requirement</span></span>| <span data-ttu-id="27d2a-888">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-888">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-889">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-889">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27d2a-890">1.7</span><span class="sxs-lookup"><span data-stu-id="27d2a-890">1.7</span></span> |
|[<span data-ttu-id="27d2a-891">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-891">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27d2a-892">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-892">ReadItem</span></span> |
|[<span data-ttu-id="27d2a-893">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-893">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="27d2a-894">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-894">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="27d2a-895">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-895">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="27d2a-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="27d2a-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="27d2a-897">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-897">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="27d2a-p149">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="27d2a-901">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="27d2a-901">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="27d2a-902">Se o Suplemento do Office estiver em execução no Outlook na Web, o método `addItemAttachmentAsync` pode anexar itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="27d2a-902">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-903">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-903">Parameters</span></span>

|<span data-ttu-id="27d2a-904">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-904">Name</span></span>|<span data-ttu-id="27d2a-905">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-905">Type</span></span>|<span data-ttu-id="27d2a-906">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-906">Attributes</span></span>|<span data-ttu-id="27d2a-907">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-907">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="27d2a-908">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-908">String</span></span>||<span data-ttu-id="27d2a-p150">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="27d2a-911">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-911">String</span></span>||<span data-ttu-id="27d2a-912">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="27d2a-912">The subject of the item to be attached.</span></span> <span data-ttu-id="27d2a-913">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="27d2a-913">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="27d2a-914">Object</span><span class="sxs-lookup"><span data-stu-id="27d2a-914">Object</span></span>|<span data-ttu-id="27d2a-915">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-915">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-916">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-916">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="27d2a-917">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-917">Object</span></span>|<span data-ttu-id="27d2a-918">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-918">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-919">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-919">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="27d2a-920">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-920">function</span></span>|<span data-ttu-id="27d2a-921">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-921">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-922">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-922">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="27d2a-923">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-923">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="27d2a-924">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="27d2a-924">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="27d2a-925">Erros</span><span class="sxs-lookup"><span data-stu-id="27d2a-925">Errors</span></span>

|<span data-ttu-id="27d2a-926">Código de erro</span><span class="sxs-lookup"><span data-stu-id="27d2a-926">Error code</span></span>|<span data-ttu-id="27d2a-927">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-927">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="27d2a-928">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="27d2a-928">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-929">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-929">Requirements</span></span>

|<span data-ttu-id="27d2a-930">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-930">Requirement</span></span>|<span data-ttu-id="27d2a-931">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-931">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-932">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-932">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-933">1.1</span><span class="sxs-lookup"><span data-stu-id="27d2a-933">1.1</span></span>|
|[<span data-ttu-id="27d2a-934">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-934">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-935">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-935">ReadWriteItem</span></span>|
|[<span data-ttu-id="27d2a-936">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-936">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-937">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-937">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-938">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-938">Example</span></span>

<span data-ttu-id="27d2a-939">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-939">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="27d2a-940">close()</span><span class="sxs-lookup"><span data-stu-id="27d2a-940">close()</span></span>

<span data-ttu-id="27d2a-941">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="27d2a-941">Closes the current item that is being composed.</span></span>

<span data-ttu-id="27d2a-p152">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-944">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="27d2a-944">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="27d2a-945">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="27d2a-945">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-946">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-946">Requirements</span></span>

|<span data-ttu-id="27d2a-947">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-947">Requirement</span></span>|<span data-ttu-id="27d2a-948">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-948">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-949">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-949">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-950">1.3</span><span class="sxs-lookup"><span data-stu-id="27d2a-950">1.3</span></span>|
|[<span data-ttu-id="27d2a-951">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-951">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-952">Restrito</span><span class="sxs-lookup"><span data-stu-id="27d2a-952">Restricted</span></span>|
|[<span data-ttu-id="27d2a-953">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-953">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-954">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-954">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="27d2a-955">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="27d2a-955">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="27d2a-956">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="27d2a-956">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-957">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="27d2a-957">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="27d2a-958">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="27d2a-958">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="27d2a-959">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="27d2a-959">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="27d2a-p153">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-963">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-963">Parameters</span></span>

|<span data-ttu-id="27d2a-964">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-964">Name</span></span>|<span data-ttu-id="27d2a-965">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-965">Type</span></span>|<span data-ttu-id="27d2a-966">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-966">Attributes</span></span>|<span data-ttu-id="27d2a-967">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-967">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="27d2a-968">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="27d2a-968">String &#124; Object</span></span>||<span data-ttu-id="27d2a-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="27d2a-971">**OU**</span><span class="sxs-lookup"><span data-stu-id="27d2a-971">**OR**</span></span><br/><span data-ttu-id="27d2a-p155">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="27d2a-974">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-974">String</span></span>|<span data-ttu-id="27d2a-975">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-975">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-p156">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="27d2a-978">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-978">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="27d2a-979">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-979">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-980">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-980">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="27d2a-981">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-981">String</span></span>||<span data-ttu-id="27d2a-p157">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="27d2a-984">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-984">String</span></span>||<span data-ttu-id="27d2a-985">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="27d2a-985">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="27d2a-986">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-986">String</span></span>||<span data-ttu-id="27d2a-p158">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="27d2a-989">Booliano</span><span class="sxs-lookup"><span data-stu-id="27d2a-989">Boolean</span></span>||<span data-ttu-id="27d2a-p159">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="27d2a-992">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-992">String</span></span>||<span data-ttu-id="27d2a-p160">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="27d2a-996">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-996">function</span></span>|<span data-ttu-id="27d2a-997">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-997">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-998">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-998">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-999">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-999">Requirements</span></span>

|<span data-ttu-id="27d2a-1000">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1000">Requirement</span></span>|<span data-ttu-id="27d2a-1001">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1002">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-1003">1.0</span></span>|
|[<span data-ttu-id="27d2a-1004">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1004">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1005">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1006">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-1006">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1007">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-1007">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="27d2a-1008">Exemplos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1008">Examples</span></span>

<span data-ttu-id="27d2a-1009">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1009">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="27d2a-1010">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1010">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="27d2a-1011">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1011">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="27d2a-1012">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1012">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="27d2a-1013">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1013">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="27d2a-1014">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1014">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="27d2a-1015">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="27d2a-1015">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="27d2a-1016">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1016">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-1017">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1017">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="27d2a-1018">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1018">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="27d2a-1019">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1019">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="27d2a-p161">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1023">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1023">Parameters</span></span>

|<span data-ttu-id="27d2a-1024">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1024">Name</span></span>|<span data-ttu-id="27d2a-1025">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1025">Type</span></span>|<span data-ttu-id="27d2a-1026">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1026">Attributes</span></span>|<span data-ttu-id="27d2a-1027">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1027">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="27d2a-1028">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="27d2a-1028">String &#124; Object</span></span>||<span data-ttu-id="27d2a-p162">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="27d2a-1031">**OU**</span><span class="sxs-lookup"><span data-stu-id="27d2a-1031">**OR**</span></span><br/><span data-ttu-id="27d2a-p163">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="27d2a-1034">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-1034">String</span></span>|<span data-ttu-id="27d2a-1035">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1035">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-p164">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="27d2a-1038">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1038">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="27d2a-1039">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1040">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1040">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="27d2a-1041">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-1041">String</span></span>||<span data-ttu-id="27d2a-p165">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="27d2a-1044">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-1044">String</span></span>||<span data-ttu-id="27d2a-1045">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1045">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="27d2a-1046">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-1046">String</span></span>||<span data-ttu-id="27d2a-p166">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="27d2a-1049">Booliano</span><span class="sxs-lookup"><span data-stu-id="27d2a-1049">Boolean</span></span>||<span data-ttu-id="27d2a-p167">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="27d2a-1052">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-1052">String</span></span>||<span data-ttu-id="27d2a-p168">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="27d2a-1056">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-1056">function</span></span>|<span data-ttu-id="27d2a-1057">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1058">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1059">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1059">Requirements</span></span>

|<span data-ttu-id="27d2a-1060">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1060">Requirement</span></span>|<span data-ttu-id="27d2a-1061">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1061">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1062">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1062">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1063">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-1063">1.0</span></span>|
|[<span data-ttu-id="27d2a-1064">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1064">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1065">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1065">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1066">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1066">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1067">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-1067">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="27d2a-1068">Exemplos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1068">Examples</span></span>

<span data-ttu-id="27d2a-1069">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1069">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="27d2a-1070">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1070">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="27d2a-1071">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1071">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="27d2a-1072">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1072">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="27d2a-1073">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1073">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="27d2a-1074">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1074">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="27d2a-1075">getAllInternetHeadersAsync ([opções], [callback])</span><span class="sxs-lookup"><span data-stu-id="27d2a-1075">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="27d2a-1076">Obtém todos os cabeçalhos de Internet da mensagem como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1076">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="27d2a-1077">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1077">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1078">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1078">Parameters</span></span>

|<span data-ttu-id="27d2a-1079">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1079">Name</span></span>|<span data-ttu-id="27d2a-1080">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1080">Type</span></span>|<span data-ttu-id="27d2a-1081">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1081">Attributes</span></span>|<span data-ttu-id="27d2a-1082">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1082">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="27d2a-1083">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1083">Object</span></span>|<span data-ttu-id="27d2a-1084">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1084">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1085">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1085">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="27d2a-1086">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1086">Object</span></span>|<span data-ttu-id="27d2a-1087">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1087">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1088">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1088">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="27d2a-1089">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-1089">function</span></span>|<span data-ttu-id="27d2a-1090">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1091">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1091">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="27d2a-1092">Com êxito, os dados de cabeçalhos de Internet são fornecidos na propriedade asyncResult. Value como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1092">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="27d2a-1093">Consulte [RFC 2183](https://tools.ietf.org/html/rfc2183) para obter as informações de formatação do valor de cadeia de caracteres retornado.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1093">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="27d2a-1094">Se a chamada falhar, a propriedade asyncResult. Error conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1094">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1095">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1095">Requirements</span></span>

|<span data-ttu-id="27d2a-1096">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1096">Requirement</span></span>|<span data-ttu-id="27d2a-1097">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1098">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1099">1,8</span><span class="sxs-lookup"><span data-stu-id="27d2a-1099">1.8</span></span>|
|[<span data-ttu-id="27d2a-1100">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1101">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1102">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1103">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-1103">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="27d2a-1104">Retorna:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1104">Returns:</span></span>

<span data-ttu-id="27d2a-1105">A Internet cabeçalhos dados como uma cadeia de caracteres formatada de acordo com a [RFC 2183](https://tools.ietf.org/html/rfc2183).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1105">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="27d2a-1106">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="27d2a-1106">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="27d2a-1107">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1107">Example</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="27d2a-1108">getAttachmentContentAsync (attachmentid, [opções], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="27d2a-1108">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="27d2a-1109">Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um `AttachmentContent` objeto.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1109">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="27d2a-1110">O `getAttachmentContentAsync` método obtém o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1110">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="27d2a-1111">Como prática recomendada, você deve usar o identificador para recuperar um anexo na mesma sessão em que o attachmentIds foi recuperado com a `getAttachmentsAsync` chamada ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="27d2a-1111">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="27d2a-1112">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1112">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="27d2a-1113">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1113">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1114">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1114">Parameters</span></span>

|<span data-ttu-id="27d2a-1115">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1115">Name</span></span>|<span data-ttu-id="27d2a-1116">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1116">Type</span></span>|<span data-ttu-id="27d2a-1117">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1117">Attributes</span></span>|<span data-ttu-id="27d2a-1118">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1118">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="27d2a-1119">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-1119">String</span></span>||<span data-ttu-id="27d2a-1120">O identificador do anexo que você deseja obter.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1120">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="27d2a-1121">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1121">Object</span></span>|<span data-ttu-id="27d2a-1122">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1122">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1123">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1123">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="27d2a-1124">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1124">Object</span></span>|<span data-ttu-id="27d2a-1125">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1125">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1126">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1126">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="27d2a-1127">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-1127">function</span></span>|<span data-ttu-id="27d2a-1128">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1129">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1129">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1130">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1130">Requirements</span></span>

|<span data-ttu-id="27d2a-1131">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1131">Requirement</span></span>|<span data-ttu-id="27d2a-1132">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1133">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1134">1,8</span><span class="sxs-lookup"><span data-stu-id="27d2a-1134">1.8</span></span>|
|[<span data-ttu-id="27d2a-1135">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1136">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1137">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-1137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1138">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-1138">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="27d2a-1139">Retorna:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1139">Returns:</span></span>

<span data-ttu-id="27d2a-1140">Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="27d2a-1140">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="27d2a-1141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1141">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="27d2a-1142">getAttachmentsAsync ([Options], [callback]) → array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="27d2a-1142">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="27d2a-1143">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1143">Gets the item's attachments as an array.</span></span> <span data-ttu-id="27d2a-1144">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1144">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1145">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1145">Parameters</span></span>

|<span data-ttu-id="27d2a-1146">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1146">Name</span></span>|<span data-ttu-id="27d2a-1147">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1147">Type</span></span>|<span data-ttu-id="27d2a-1148">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1148">Attributes</span></span>|<span data-ttu-id="27d2a-1149">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1149">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="27d2a-1150">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1150">Object</span></span>|<span data-ttu-id="27d2a-1151">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1151">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1152">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1152">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="27d2a-1153">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1153">Object</span></span>|<span data-ttu-id="27d2a-1154">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1155">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1155">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="27d2a-1156">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-1156">function</span></span>|<span data-ttu-id="27d2a-1157">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1158">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1158">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1159">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1159">Requirements</span></span>

|<span data-ttu-id="27d2a-1160">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1160">Requirement</span></span>|<span data-ttu-id="27d2a-1161">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1162">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1163">1,8</span><span class="sxs-lookup"><span data-stu-id="27d2a-1163">1.8</span></span>|
|[<span data-ttu-id="27d2a-1164">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1165">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1165">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1166">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1167">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-1167">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="27d2a-1168">Retorna:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1168">Returns:</span></span>

<span data-ttu-id="27d2a-1169">Tipo: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="27d2a-1169">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="27d2a-1170">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1170">Example</span></span>

<span data-ttu-id="27d2a-1171">O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1171">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="27d2a-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="27d2a-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="27d2a-1173">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1173">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-1174">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1174">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-1175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1175">Requirements</span></span>

|<span data-ttu-id="27d2a-1176">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1176">Requirement</span></span>|<span data-ttu-id="27d2a-1177">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1177">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1179">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-1179">1.0</span></span>|
|[<span data-ttu-id="27d2a-1180">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1180">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1181">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1182">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1182">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1183">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-1183">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="27d2a-1184">Retorna:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1184">Returns:</span></span>

<span data-ttu-id="27d2a-1185">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="27d2a-1185">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="27d2a-1186">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1186">Example</span></span>

<span data-ttu-id="27d2a-1187">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1187">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="27d2a-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="27d2a-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="27d2a-1189">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1189">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-1190">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1190">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1191">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1191">Parameters</span></span>

|<span data-ttu-id="27d2a-1192">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1192">Name</span></span>|<span data-ttu-id="27d2a-1193">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1193">Type</span></span>|<span data-ttu-id="27d2a-1194">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1194">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="27d2a-1195">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="27d2a-1195">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="27d2a-1196">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1196">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1197">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1197">Requirements</span></span>

|<span data-ttu-id="27d2a-1198">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1198">Requirement</span></span>|<span data-ttu-id="27d2a-1199">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1199">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1200">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1201">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-1201">1.0</span></span>|
|[<span data-ttu-id="27d2a-1202">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1203">Restrito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1203">Restricted</span></span>|
|[<span data-ttu-id="27d2a-1204">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1205">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-1205">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="27d2a-1206">Retorna:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1206">Returns:</span></span>

<span data-ttu-id="27d2a-1207">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1207">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="27d2a-1208">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1208">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="27d2a-1209">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1209">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="27d2a-1210">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1210">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="27d2a-1211">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="27d2a-1211">Value of `entityType`</span></span>|<span data-ttu-id="27d2a-1212">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="27d2a-1212">Type of objects in returned array</span></span>|<span data-ttu-id="27d2a-1213">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="27d2a-1213">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="27d2a-1214">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-1214">String</span></span>|<span data-ttu-id="27d2a-1215">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="27d2a-1215">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="27d2a-1216">Contato</span><span class="sxs-lookup"><span data-stu-id="27d2a-1216">Contact</span></span>|<span data-ttu-id="27d2a-1217">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="27d2a-1217">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="27d2a-1218">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-1218">String</span></span>|<span data-ttu-id="27d2a-1219">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="27d2a-1219">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="27d2a-1220">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="27d2a-1220">MeetingSuggestion</span></span>|<span data-ttu-id="27d2a-1221">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="27d2a-1221">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="27d2a-1222">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="27d2a-1222">PhoneNumber</span></span>|<span data-ttu-id="27d2a-1223">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="27d2a-1223">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="27d2a-1224">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="27d2a-1224">TaskSuggestion</span></span>|<span data-ttu-id="27d2a-1225">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="27d2a-1225">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="27d2a-1226">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-1226">String</span></span>|<span data-ttu-id="27d2a-1227">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="27d2a-1227">**Restricted**</span></span>|

<span data-ttu-id="27d2a-1228">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="27d2a-1228">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="27d2a-1229">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1229">Example</span></span>

<span data-ttu-id="27d2a-1230">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1230">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="27d2a-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="27d2a-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="27d2a-1232">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1232">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-1233">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1233">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="27d2a-1234">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1234">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1235">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1235">Parameters</span></span>

|<span data-ttu-id="27d2a-1236">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1236">Name</span></span>|<span data-ttu-id="27d2a-1237">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1237">Type</span></span>|<span data-ttu-id="27d2a-1238">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1238">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="27d2a-1239">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="27d2a-1239">String</span></span>|<span data-ttu-id="27d2a-1240">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1240">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1241">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1241">Requirements</span></span>

|<span data-ttu-id="27d2a-1242">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1242">Requirement</span></span>|<span data-ttu-id="27d2a-1243">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1243">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1244">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1244">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1245">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-1245">1.0</span></span>|
|[<span data-ttu-id="27d2a-1246">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1246">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1247">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1247">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1248">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1248">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1249">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-1249">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="27d2a-1250">Retorna:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1250">Returns:</span></span>

<span data-ttu-id="27d2a-p174">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="27d2a-1253">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="27d2a-1253">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="27d2a-1254">getInitializationContextAsync ([opções], [callback])</span><span class="sxs-lookup"><span data-stu-id="27d2a-1254">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="27d2a-1255">Obtém dados de inicialização passados quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1255">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-1256">Este método só é compatível com o Outlook 2016 ou posterior no Windows (clique para executar versões posteriores a 16.0.8413.1000) e Outlook na Web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1256">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1257">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1257">Parameters</span></span>

|<span data-ttu-id="27d2a-1258">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1258">Name</span></span>|<span data-ttu-id="27d2a-1259">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1259">Type</span></span>|<span data-ttu-id="27d2a-1260">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1260">Attributes</span></span>|<span data-ttu-id="27d2a-1261">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1261">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="27d2a-1262">Object</span><span class="sxs-lookup"><span data-stu-id="27d2a-1262">Object</span></span>|<span data-ttu-id="27d2a-1263">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1263">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1264">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1264">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="27d2a-1265">Object</span><span class="sxs-lookup"><span data-stu-id="27d2a-1265">Object</span></span>|<span data-ttu-id="27d2a-1266">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1266">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1267">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1267">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="27d2a-1268">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-1268">function</span></span>|<span data-ttu-id="27d2a-1269">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1269">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1270">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1270">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="27d2a-1271">Com êxito, os dados de inicialização são fornecidos na `asyncResult.value` Propriedade como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1271">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="27d2a-1272">Se não houver nenhum contexto de inicialização, `asyncResult` o objeto conterá `Error` um objeto com `code` sua propriedade definida `9020` como e `name` sua propriedade definida `GenericResponseError`como.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1272">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1273">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1273">Requirements</span></span>

|<span data-ttu-id="27d2a-1274">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1274">Requirement</span></span>|<span data-ttu-id="27d2a-1275">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1275">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1276">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1277">Visualização</span><span class="sxs-lookup"><span data-stu-id="27d2a-1277">Preview</span></span>|
|[<span data-ttu-id="27d2a-1278">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1279">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1280">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1281">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-1281">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-1282">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1282">Example</span></span>

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

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="27d2a-1283">getItemIdAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="27d2a-1283">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="27d2a-1284">Obtém de forma assíncrona a ID de um item salvo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1284">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="27d2a-1285">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1285">Compose mode only.</span></span>

<span data-ttu-id="27d2a-1286">Quando invocado, este método retorna a ID do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1286">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-1287">Se seu suplemento chamar `getItemIdAsync` um item no modo de redação (por exemplo, para `itemId` usar com o EWS ou a API REST), lembre-se de que, quando o Outlook estiver no modo cache, pode levar algum tempo para que o item seja sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1287">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="27d2a-1288">Até que o item seja sincronizado, `itemId` o não é reconhecido e usado retorna um erro.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1288">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1289">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1289">Parameters</span></span>

|<span data-ttu-id="27d2a-1290">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1290">Name</span></span>|<span data-ttu-id="27d2a-1291">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1291">Type</span></span>|<span data-ttu-id="27d2a-1292">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1292">Attributes</span></span>|<span data-ttu-id="27d2a-1293">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1293">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="27d2a-1294">Object</span><span class="sxs-lookup"><span data-stu-id="27d2a-1294">Object</span></span>|<span data-ttu-id="27d2a-1295">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1295">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1296">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1296">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="27d2a-1297">Object</span><span class="sxs-lookup"><span data-stu-id="27d2a-1297">Object</span></span>|<span data-ttu-id="27d2a-1298">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1298">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1299">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1299">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="27d2a-1300">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-1300">function</span></span>||<span data-ttu-id="27d2a-1301">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1301">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="27d2a-1302">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1302">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="27d2a-1303">Erros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1303">Errors</span></span>

|<span data-ttu-id="27d2a-1304">Código de erro</span><span class="sxs-lookup"><span data-stu-id="27d2a-1304">Error code</span></span>|<span data-ttu-id="27d2a-1305">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1305">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="27d2a-1306">A ID não pode ser recuperada até que o item seja salvo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1306">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1307">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1307">Requirements</span></span>

|<span data-ttu-id="27d2a-1308">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1308">Requirement</span></span>|<span data-ttu-id="27d2a-1309">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1309">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1310">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1310">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1311">1,8</span><span class="sxs-lookup"><span data-stu-id="27d2a-1311">1.8</span></span>|
|[<span data-ttu-id="27d2a-1312">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1312">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1313">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1314">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1314">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1315">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-1315">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="27d2a-1316">Exemplos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1316">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="27d2a-1317">O exemplo a seguir mostra a estrutura do `result` parâmetro que é passado para a função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1317">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="27d2a-1318">A `value` propriedade contém a ID do item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1318">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="27d2a-1319">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="27d2a-1319">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="27d2a-1320">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1320">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-1321">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1321">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="27d2a-p178">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="27d2a-1325">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1325">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="27d2a-1326">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1326">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="27d2a-p179">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-1330">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1330">Requirements</span></span>

|<span data-ttu-id="27d2a-1331">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1331">Requirement</span></span>|<span data-ttu-id="27d2a-1332">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1332">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1333">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1334">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-1334">1.0</span></span>|
|[<span data-ttu-id="27d2a-1335">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1336">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1337">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1338">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-1338">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="27d2a-1339">Retorna:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1339">Returns:</span></span>

<span data-ttu-id="27d2a-p180">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="27d2a-1342">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="27d2a-1342">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="27d2a-1343">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1343">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="27d2a-1344">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1344">Example</span></span>

<span data-ttu-id="27d2a-1345">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1345">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="27d2a-1346">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="27d2a-1346">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="27d2a-1347">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1347">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-1348">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="27d2a-1349">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1349">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="27d2a-p181">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1352">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1352">Parameters</span></span>

|<span data-ttu-id="27d2a-1353">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1353">Name</span></span>|<span data-ttu-id="27d2a-1354">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1354">Type</span></span>|<span data-ttu-id="27d2a-1355">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1355">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="27d2a-1356">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-1356">String</span></span>|<span data-ttu-id="27d2a-1357">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1357">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1358">Requirements</span></span>

|<span data-ttu-id="27d2a-1359">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1359">Requirement</span></span>|<span data-ttu-id="27d2a-1360">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1360">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1361">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1362">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-1362">1.0</span></span>|
|[<span data-ttu-id="27d2a-1363">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1364">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1365">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1366">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-1366">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="27d2a-1367">Retorna:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1367">Returns:</span></span>

<span data-ttu-id="27d2a-1368">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1368">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="27d2a-1369">Tipo: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="27d2a-1369">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="27d2a-1370">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1370">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="27d2a-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="27d2a-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="27d2a-1372">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1372">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="27d2a-1373">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retornará uma cadeia de caracteres vazia para os dados selecionados.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1373">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.</span></span> <span data-ttu-id="27d2a-1374">Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1374">If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-1375">No Outlook na Web, o método retorna a cadeia de caracteres “null” se nenhum texto for selecionado, mas o cursor estiver no corpo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1375">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="27d2a-1376">Para verificar essa situação, confira o exemplo mais adiante nesta seção.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1376">To check for this situation, see the example later in this section.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1377">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1377">Parameters</span></span>

|<span data-ttu-id="27d2a-1378">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1378">Name</span></span>|<span data-ttu-id="27d2a-1379">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1379">Type</span></span>|<span data-ttu-id="27d2a-1380">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1380">Attributes</span></span>|<span data-ttu-id="27d2a-1381">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1381">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="27d2a-1382">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="27d2a-1382">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="27d2a-p184">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p184">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="27d2a-1386">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1386">Object</span></span>|<span data-ttu-id="27d2a-1387">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1387">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1388">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1388">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="27d2a-1389">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1389">Object</span></span>|<span data-ttu-id="27d2a-1390">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1390">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1391">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1391">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="27d2a-1392">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-1392">function</span></span>||<span data-ttu-id="27d2a-1393">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1393">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="27d2a-1394">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1394">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="27d2a-1395">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1395">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1396">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1396">Requirements</span></span>

|<span data-ttu-id="27d2a-1397">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1397">Requirement</span></span>|<span data-ttu-id="27d2a-1398">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1398">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1399">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1400">1.2</span><span class="sxs-lookup"><span data-stu-id="27d2a-1400">1.2</span></span>|
|[<span data-ttu-id="27d2a-1401">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1401">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1402">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1403">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1403">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1404">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-1404">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="27d2a-1405">Retorna:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1405">Returns:</span></span>

<span data-ttu-id="27d2a-1406">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1406">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="27d2a-1407">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="27d2a-1407">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="27d2a-1408">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1408">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="27d2a-1409">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="27d2a-1409">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="27d2a-1410">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1410">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="27d2a-1411">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1411">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-1412">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1412">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-1413">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1413">Requirements</span></span>

|<span data-ttu-id="27d2a-1414">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1414">Requirement</span></span>|<span data-ttu-id="27d2a-1415">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1415">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1416">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1417">1.6</span><span class="sxs-lookup"><span data-stu-id="27d2a-1417">1.6</span></span>|
|[<span data-ttu-id="27d2a-1418">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1418">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1419">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1419">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1420">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1420">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1421">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-1421">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="27d2a-1422">Retorna:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1422">Returns:</span></span>

<span data-ttu-id="27d2a-1423">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="27d2a-1423">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="27d2a-1424">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1424">Example</span></span>

<span data-ttu-id="27d2a-1425">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1425">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="27d2a-1426">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="27d2a-1426">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="27d2a-p187">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="27d2a-p187">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-1429">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1429">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="27d2a-p188">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p188">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="27d2a-1433">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1433">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="27d2a-1434">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1434">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="27d2a-p189">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p189">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="27d2a-1438">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1438">Requirements</span></span>

|<span data-ttu-id="27d2a-1439">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1439">Requirement</span></span>|<span data-ttu-id="27d2a-1440">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1440">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1441">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1442">1.6</span><span class="sxs-lookup"><span data-stu-id="27d2a-1442">1.6</span></span>|
|[<span data-ttu-id="27d2a-1443">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1444">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1445">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1446">Read</span><span class="sxs-lookup"><span data-stu-id="27d2a-1446">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="27d2a-1447">Retorna:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1447">Returns:</span></span>

<span data-ttu-id="27d2a-p190">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p190">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="27d2a-1450">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1450">Example</span></span>

<span data-ttu-id="27d2a-1451">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1451">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="27d2a-1452">getSharedPropertiesAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="27d2a-1452">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="27d2a-1453">Obtém as propriedades do compromisso ou da mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1453">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1454">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1454">Parameters</span></span>

|<span data-ttu-id="27d2a-1455">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1455">Name</span></span>|<span data-ttu-id="27d2a-1456">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1456">Type</span></span>|<span data-ttu-id="27d2a-1457">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1457">Attributes</span></span>|<span data-ttu-id="27d2a-1458">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1458">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="27d2a-1459">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1459">Object</span></span>|<span data-ttu-id="27d2a-1460">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1460">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1461">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1461">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="27d2a-1462">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1462">Object</span></span>|<span data-ttu-id="27d2a-1463">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1463">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1464">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1464">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="27d2a-1465">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-1465">function</span></span>||<span data-ttu-id="27d2a-1466">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1466">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="27d2a-1467">As propriedades compartilhadas são fornecidas [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) como um objeto `asyncResult.value` na propriedade.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1467">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="27d2a-1468">Este objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1468">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1469">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1469">Requirements</span></span>

|<span data-ttu-id="27d2a-1470">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1470">Requirement</span></span>|<span data-ttu-id="27d2a-1471">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1471">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1472">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1473">1,8</span><span class="sxs-lookup"><span data-stu-id="27d2a-1473">1.8</span></span>|
|[<span data-ttu-id="27d2a-1474">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1475">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1476">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-1476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1477">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-1477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-1478">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1478">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="27d2a-1479">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="27d2a-1479">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="27d2a-1480">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1480">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="27d2a-p192">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p192">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1484">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1484">Parameters</span></span>

|<span data-ttu-id="27d2a-1485">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1485">Name</span></span>|<span data-ttu-id="27d2a-1486">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1486">Type</span></span>|<span data-ttu-id="27d2a-1487">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1487">Attributes</span></span>|<span data-ttu-id="27d2a-1488">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1488">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="27d2a-1489">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-1489">function</span></span>||<span data-ttu-id="27d2a-1490">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1490">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="27d2a-1491">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1491">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="27d2a-1492">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1492">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="27d2a-1493">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1493">Object</span></span>|<span data-ttu-id="27d2a-1494">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1494">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1495">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1495">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="27d2a-1496">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1496">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1497">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1497">Requirements</span></span>

|<span data-ttu-id="27d2a-1498">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1498">Requirement</span></span>|<span data-ttu-id="27d2a-1499">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1499">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1500">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1500">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1501">1.0</span><span class="sxs-lookup"><span data-stu-id="27d2a-1501">1.0</span></span>|
|[<span data-ttu-id="27d2a-1502">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1502">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1503">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1503">ReadItem</span></span>|
|[<span data-ttu-id="27d2a-1504">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-1504">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1505">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-1505">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-1506">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1506">Example</span></span>

<span data-ttu-id="27d2a-p195">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p195">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="27d2a-1510">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="27d2a-1510">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="27d2a-1511">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1511">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="27d2a-1512">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1512">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="27d2a-1513">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1513">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="27d2a-1514">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1514">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="27d2a-1515">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1515">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1516">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1516">Parameters</span></span>

|<span data-ttu-id="27d2a-1517">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1517">Name</span></span>|<span data-ttu-id="27d2a-1518">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1518">Type</span></span>|<span data-ttu-id="27d2a-1519">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1519">Attributes</span></span>|<span data-ttu-id="27d2a-1520">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1520">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="27d2a-1521">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-1521">String</span></span>||<span data-ttu-id="27d2a-1522">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1522">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="27d2a-1523">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1523">Object</span></span>|<span data-ttu-id="27d2a-1524">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1524">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1525">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1525">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="27d2a-1526">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1526">Object</span></span>|<span data-ttu-id="27d2a-1527">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1527">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1528">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1528">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="27d2a-1529">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-1529">function</span></span>|<span data-ttu-id="27d2a-1530">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1530">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1531">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1531">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="27d2a-1532">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1532">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="27d2a-1533">Erros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1533">Errors</span></span>

|<span data-ttu-id="27d2a-1534">Código de erro</span><span class="sxs-lookup"><span data-stu-id="27d2a-1534">Error code</span></span>|<span data-ttu-id="27d2a-1535">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1535">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="27d2a-1536">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1536">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1537">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1537">Requirements</span></span>

|<span data-ttu-id="27d2a-1538">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1538">Requirement</span></span>|<span data-ttu-id="27d2a-1539">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1539">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1540">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1541">1.1</span><span class="sxs-lookup"><span data-stu-id="27d2a-1541">1.1</span></span>|
|[<span data-ttu-id="27d2a-1542">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1542">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1543">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1543">ReadWriteItem</span></span>|
|[<span data-ttu-id="27d2a-1544">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1545">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-1545">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-1546">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1546">Example</span></span>

<span data-ttu-id="27d2a-1547">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1547">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="27d2a-1548">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="27d2a-1548">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="27d2a-1549">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1549">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="27d2a-1550">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1550">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1551">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1551">Parameters</span></span>

| <span data-ttu-id="27d2a-1552">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1552">Name</span></span> | <span data-ttu-id="27d2a-1553">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1553">Type</span></span> | <span data-ttu-id="27d2a-1554">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1554">Attributes</span></span> | <span data-ttu-id="27d2a-1555">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1555">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="27d2a-1556">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="27d2a-1556">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="27d2a-1557">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1557">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="27d2a-1558">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1558">Object</span></span> | <span data-ttu-id="27d2a-1559">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1559">&lt;optional&gt;</span></span> | <span data-ttu-id="27d2a-1560">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1560">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="27d2a-1561">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1561">Object</span></span> | <span data-ttu-id="27d2a-1562">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1562">&lt;optional&gt;</span></span> | <span data-ttu-id="27d2a-1563">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1563">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="27d2a-1564">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-1564">function</span></span>| <span data-ttu-id="27d2a-1565">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1565">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1566">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1566">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1567">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1567">Requirements</span></span>

|<span data-ttu-id="27d2a-1568">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1568">Requirement</span></span>| <span data-ttu-id="27d2a-1569">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1569">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1570">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27d2a-1571">1.7</span><span class="sxs-lookup"><span data-stu-id="27d2a-1571">1.7</span></span> |
|[<span data-ttu-id="27d2a-1572">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1572">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27d2a-1573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1573">ReadItem</span></span> |
|[<span data-ttu-id="27d2a-1574">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="27d2a-1574">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="27d2a-1575">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="27d2a-1575">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="27d2a-1576">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="27d2a-1576">saveAsync([options], callback)</span></span>

<span data-ttu-id="27d2a-1577">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1577">Asynchronously saves an item.</span></span>

<span data-ttu-id="27d2a-1578">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1578">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="27d2a-1579">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1579">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="27d2a-1580">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1580">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-1581">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1581">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="27d2a-1582">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1582">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="27d2a-p199">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p199">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="27d2a-1586">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="27d2a-1586">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="27d2a-1587">O Outlook no Mac não dá suporte ao salvamento de reuniões.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1587">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="27d2a-1588">O método `saveAsync` falha quando chamado a partir de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1588">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="27d2a-1589">Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1589">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="27d2a-1590">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1590">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1591">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1591">Parameters</span></span>

|<span data-ttu-id="27d2a-1592">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1592">Name</span></span>|<span data-ttu-id="27d2a-1593">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1593">Type</span></span>|<span data-ttu-id="27d2a-1594">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1594">Attributes</span></span>|<span data-ttu-id="27d2a-1595">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1595">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="27d2a-1596">Object</span><span class="sxs-lookup"><span data-stu-id="27d2a-1596">Object</span></span>|<span data-ttu-id="27d2a-1597">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1597">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1598">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1598">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="27d2a-1599">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1599">Object</span></span>|<span data-ttu-id="27d2a-1600">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1600">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1601">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1601">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="27d2a-1602">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-1602">function</span></span>||<span data-ttu-id="27d2a-1603">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1603">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="27d2a-1604">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1604">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1605">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1605">Requirements</span></span>

|<span data-ttu-id="27d2a-1606">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1606">Requirement</span></span>|<span data-ttu-id="27d2a-1607">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1607">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1608">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1609">1.3</span><span class="sxs-lookup"><span data-stu-id="27d2a-1609">1.3</span></span>|
|[<span data-ttu-id="27d2a-1610">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1611">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1611">ReadWriteItem</span></span>|
|[<span data-ttu-id="27d2a-1612">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1613">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-1613">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="27d2a-1614">Exemplos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1614">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="27d2a-p201">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p201">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="27d2a-1617">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="27d2a-1617">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="27d2a-1618">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1618">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="27d2a-p202">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p202">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27d2a-1622">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="27d2a-1622">Parameters</span></span>

|<span data-ttu-id="27d2a-1623">Nome</span><span class="sxs-lookup"><span data-stu-id="27d2a-1623">Name</span></span>|<span data-ttu-id="27d2a-1624">Tipo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1624">Type</span></span>|<span data-ttu-id="27d2a-1625">Atributos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1625">Attributes</span></span>|<span data-ttu-id="27d2a-1626">Descrição</span><span class="sxs-lookup"><span data-stu-id="27d2a-1626">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="27d2a-1627">String</span><span class="sxs-lookup"><span data-stu-id="27d2a-1627">String</span></span>||<span data-ttu-id="27d2a-p203">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-p203">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="27d2a-1631">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1631">Object</span></span>|<span data-ttu-id="27d2a-1632">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1632">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1633">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1633">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="27d2a-1634">Objeto</span><span class="sxs-lookup"><span data-stu-id="27d2a-1634">Object</span></span>|<span data-ttu-id="27d2a-1635">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1635">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1636">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1636">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="27d2a-1637">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="27d2a-1637">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="27d2a-1638">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="27d2a-1638">&lt;optional&gt;</span></span>|<span data-ttu-id="27d2a-1639">Se `text`, o estilo atual é aplicado nos clientes do Outlook na Web e do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1639">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="27d2a-1640">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1640">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="27d2a-1641">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook na Web e o estilo padrão é aplicado nos clientes do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1641">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="27d2a-1642">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1642">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="27d2a-1643">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="27d2a-1643">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="27d2a-1644">function</span><span class="sxs-lookup"><span data-stu-id="27d2a-1644">function</span></span>||<span data-ttu-id="27d2a-1645">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27d2a-1645">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27d2a-1646">Requisitos</span><span class="sxs-lookup"><span data-stu-id="27d2a-1646">Requirements</span></span>

|<span data-ttu-id="27d2a-1647">Requisito</span><span class="sxs-lookup"><span data-stu-id="27d2a-1647">Requirement</span></span>|<span data-ttu-id="27d2a-1648">Valor</span><span class="sxs-lookup"><span data-stu-id="27d2a-1648">Value</span></span>|
|---|---|
|[<span data-ttu-id="27d2a-1649">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="27d2a-1649">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="27d2a-1650">1.2</span><span class="sxs-lookup"><span data-stu-id="27d2a-1650">1.2</span></span>|
|[<span data-ttu-id="27d2a-1651">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1651">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="27d2a-1652">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="27d2a-1652">ReadWriteItem</span></span>|
|[<span data-ttu-id="27d2a-1653">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="27d2a-1653">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="27d2a-1654">Escrever</span><span class="sxs-lookup"><span data-stu-id="27d2a-1654">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="27d2a-1655">Exemplo</span><span class="sxs-lookup"><span data-stu-id="27d2a-1655">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
