---
title: Office.Context.Mailbox.item - conjunto de requisições de visualização
description: ''
ms.date: 01/16/2019
localization_priority: Normal
ms.openlocfilehash: b4b2ec9c735270d9b1bfca3d1c24ef6b0f1ca1cb
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389596"
---
# <a name="item"></a><span data-ttu-id="a85e4-102">item</span><span class="sxs-lookup"><span data-stu-id="a85e4-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="a85e4-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="a85e4-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="a85e4-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="a85e4-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-106">Requirements</span></span>

|<span data-ttu-id="a85e4-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-107">Requirement</span></span>|<span data-ttu-id="a85e4-108">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-110">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-110">1.0</span></span>|
|[<span data-ttu-id="a85e4-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="a85e4-112">Restricted</span></span>|
|[<span data-ttu-id="a85e4-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-114">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a85e4-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="a85e4-115">Members and methods</span></span>

| <span data-ttu-id="a85e4-116">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-116">Member</span></span> | <span data-ttu-id="a85e4-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a85e4-118">attachments</span><span class="sxs-lookup"><span data-stu-id="a85e4-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="a85e4-119">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-119">Member</span></span> |
| [<span data-ttu-id="a85e4-120">bcc</span><span class="sxs-lookup"><span data-stu-id="a85e4-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a85e4-121">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-121">Member</span></span> |
| [<span data-ttu-id="a85e4-122">body</span><span class="sxs-lookup"><span data-stu-id="a85e4-122">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="a85e4-123">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-123">Member</span></span> |
| [<span data-ttu-id="a85e4-124">cc</span><span class="sxs-lookup"><span data-stu-id="a85e4-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a85e4-125">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-125">Member</span></span> |
| [<span data-ttu-id="a85e4-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="a85e4-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="a85e4-127">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-127">Member</span></span> |
| [<span data-ttu-id="a85e4-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="a85e4-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="a85e4-129">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-129">Member</span></span> |
| [<span data-ttu-id="a85e4-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="a85e4-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="a85e4-131">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-131">Member</span></span> |
| [<span data-ttu-id="a85e4-132">end</span><span class="sxs-lookup"><span data-stu-id="a85e4-132">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="a85e4-133">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-133">Member</span></span> |
| [<span data-ttu-id="a85e4-134">from</span><span class="sxs-lookup"><span data-stu-id="a85e4-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="a85e4-135">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-135">Member</span></span> |
| [<span data-ttu-id="a85e4-136">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="a85e4-136">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="a85e4-137">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-137">Member</span></span> |
| [<span data-ttu-id="a85e4-138">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="a85e4-138">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="a85e4-139">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-139">Member</span></span> |
| [<span data-ttu-id="a85e4-140">itemClass</span><span class="sxs-lookup"><span data-stu-id="a85e4-140">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="a85e4-141">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-141">Member</span></span> |
| [<span data-ttu-id="a85e4-142">itemId</span><span class="sxs-lookup"><span data-stu-id="a85e4-142">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="a85e4-143">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-143">Member</span></span> |
| [<span data-ttu-id="a85e4-144">itemType</span><span class="sxs-lookup"><span data-stu-id="a85e4-144">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="a85e4-145">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-145">Member</span></span> |
| [<span data-ttu-id="a85e4-146">location</span><span class="sxs-lookup"><span data-stu-id="a85e4-146">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="a85e4-147">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-147">Member</span></span> |
| [<span data-ttu-id="a85e4-148">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="a85e4-148">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="a85e4-149">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-149">Member</span></span> |
| [<span data-ttu-id="a85e4-150">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="a85e4-150">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="a85e4-151">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-151">Member</span></span> |
| [<span data-ttu-id="a85e4-152">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="a85e4-152">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a85e4-153">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-153">Member</span></span> |
| [<span data-ttu-id="a85e4-154">organizer</span><span class="sxs-lookup"><span data-stu-id="a85e4-154">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="a85e4-155">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-155">Member</span></span> |
| [<span data-ttu-id="a85e4-156">recurrence</span><span class="sxs-lookup"><span data-stu-id="a85e4-156">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="a85e4-157">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-157">Member</span></span> |
| [<span data-ttu-id="a85e4-158">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="a85e4-158">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a85e4-159">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-159">Member</span></span> |
| [<span data-ttu-id="a85e4-160">sender</span><span class="sxs-lookup"><span data-stu-id="a85e4-160">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="a85e4-161">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-161">Member</span></span> |
| [<span data-ttu-id="a85e4-162">seriesId</span><span class="sxs-lookup"><span data-stu-id="a85e4-162">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="a85e4-163">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-163">Member</span></span> |
| [<span data-ttu-id="a85e4-164">start</span><span class="sxs-lookup"><span data-stu-id="a85e4-164">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="a85e4-165">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-165">Member</span></span> |
| [<span data-ttu-id="a85e4-166">subject</span><span class="sxs-lookup"><span data-stu-id="a85e4-166">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="a85e4-167">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-167">Member</span></span> |
| [<span data-ttu-id="a85e4-168">to</span><span class="sxs-lookup"><span data-stu-id="a85e4-168">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a85e4-169">Membro</span><span class="sxs-lookup"><span data-stu-id="a85e4-169">Member</span></span> |
| [<span data-ttu-id="a85e4-170">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a85e4-170">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="a85e4-171">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-171">Method</span></span> |
| [<span data-ttu-id="a85e4-172">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="a85e4-172">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="a85e4-173">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-173">Method</span></span> |
| [<span data-ttu-id="a85e4-174">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a85e4-174">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a85e4-175">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-175">Method</span></span> |
| [<span data-ttu-id="a85e4-176">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a85e4-176">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="a85e4-177">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-177">Method</span></span> |
| [<span data-ttu-id="a85e4-178">close</span><span class="sxs-lookup"><span data-stu-id="a85e4-178">close</span></span>](#close) | <span data-ttu-id="a85e4-179">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-179">Method</span></span> |
| [<span data-ttu-id="a85e4-180">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="a85e4-180">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="a85e4-181">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-181">Method</span></span> |
| [<span data-ttu-id="a85e4-182">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="a85e4-182">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="a85e4-183">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-183">Method</span></span> |
| [<span data-ttu-id="a85e4-184">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="a85e4-184">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="a85e4-185">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-185">Method</span></span> |
| [<span data-ttu-id="a85e4-186">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="a85e4-186">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="a85e4-187">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-187">Method</span></span> |
| [<span data-ttu-id="a85e4-188">getEntities</span><span class="sxs-lookup"><span data-stu-id="a85e4-188">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="a85e4-189">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-189">Method</span></span> |
| [<span data-ttu-id="a85e4-190">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="a85e4-190">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="a85e4-191">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-191">Method</span></span> |
| [<span data-ttu-id="a85e4-192">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="a85e4-192">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="a85e4-193">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-193">Method</span></span> |
| [<span data-ttu-id="a85e4-194">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="a85e4-194">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="a85e4-195">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-195">Method</span></span> |
| [<span data-ttu-id="a85e4-196">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a85e4-196">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="a85e4-197">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-197">Method</span></span> |
| [<span data-ttu-id="a85e4-198">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="a85e4-198">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="a85e4-199">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-199">Method</span></span> |
| [<span data-ttu-id="a85e4-200">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a85e4-200">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="a85e4-201">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-201">Method</span></span> |
| [<span data-ttu-id="a85e4-202">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="a85e4-202">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="a85e4-203">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-203">Method</span></span> |
| [<span data-ttu-id="a85e4-204">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a85e4-204">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="a85e4-205">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-205">Method</span></span> |
| [<span data-ttu-id="a85e4-206">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a85e4-206">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="a85e4-207">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-207">Method</span></span> |
| [<span data-ttu-id="a85e4-208">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a85e4-208">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="a85e4-209">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-209">Method</span></span> |
| [<span data-ttu-id="a85e4-210">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a85e4-210">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="a85e4-211">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-211">Method</span></span> |
| [<span data-ttu-id="a85e4-212">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a85e4-212">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="a85e4-213">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-213">Method</span></span> |
| [<span data-ttu-id="a85e4-214">saveAsync</span><span class="sxs-lookup"><span data-stu-id="a85e4-214">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="a85e4-215">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-215">Method</span></span> |
| [<span data-ttu-id="a85e4-216">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a85e4-216">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="a85e4-217">Método</span><span class="sxs-lookup"><span data-stu-id="a85e4-217">Method</span></span> |

### <a name="example"></a><span data-ttu-id="a85e4-218">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-218">Example</span></span>

<span data-ttu-id="a85e4-219">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="a85e4-219">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="a85e4-220">Membros</span><span class="sxs-lookup"><span data-stu-id="a85e4-220">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="a85e4-221">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a85e4-221">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="a85e4-222">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="a85e4-222">Gets the item's attachments as an array.</span></span> <span data-ttu-id="a85e4-223">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a85e4-223">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-224">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="a85e4-224">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="a85e4-225">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="a85e4-225">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-226">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-226">Type:</span></span>

*   <span data-ttu-id="a85e4-227">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a85e4-227">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-228">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-228">Requirements</span></span>

|<span data-ttu-id="a85e4-229">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-229">Requirement</span></span>|<span data-ttu-id="a85e4-230">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-231">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-231">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-232">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-232">1.0</span></span>|
|[<span data-ttu-id="a85e4-233">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-233">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-234">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-234">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-235">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-235">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-236">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-236">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-237">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-237">Example</span></span>

<span data-ttu-id="a85e4-238">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="a85e4-238">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a85e4-239">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a85e4-239">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a85e4-240">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a85e4-240">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="a85e4-241">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="a85e4-241">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-242">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-242">Type:</span></span>

*   [<span data-ttu-id="a85e4-243">Destinatários</span><span class="sxs-lookup"><span data-stu-id="a85e4-243">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="a85e4-244">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-244">Requirements</span></span>

|<span data-ttu-id="a85e4-245">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-245">Requirement</span></span>|<span data-ttu-id="a85e4-246">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-247">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-248">1.1</span><span class="sxs-lookup"><span data-stu-id="a85e4-248">1.1</span></span>|
|[<span data-ttu-id="a85e4-249">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-250">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-251">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-252">Escrever</span><span class="sxs-lookup"><span data-stu-id="a85e4-252">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-253">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-253">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="a85e4-254">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="a85e4-254">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="a85e4-255">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="a85e4-255">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-256">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-256">Type:</span></span>

*   [<span data-ttu-id="a85e4-257">Corpo</span><span class="sxs-lookup"><span data-stu-id="a85e4-257">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="a85e4-258">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-258">Requirements</span></span>

|<span data-ttu-id="a85e4-259">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-259">Requirement</span></span>|<span data-ttu-id="a85e4-260">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-261">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-261">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-262">1.1</span><span class="sxs-lookup"><span data-stu-id="a85e4-262">1.1</span></span>|
|[<span data-ttu-id="a85e4-263">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-263">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-264">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-264">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-265">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-265">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-266">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-266">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a85e4-267">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a85e4-267">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a85e4-268">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a85e4-268">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="a85e4-269">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a85e4-269">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a85e4-270">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-270">Read mode</span></span>

<span data-ttu-id="a85e4-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a85e4-273">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a85e4-273">Compose mode</span></span>

<span data-ttu-id="a85e4-274">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="a85e4-274">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-275">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-275">Type:</span></span>

*   <span data-ttu-id="a85e4-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a85e4-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-277">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-277">Requirements</span></span>

|<span data-ttu-id="a85e4-278">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-278">Requirement</span></span>|<span data-ttu-id="a85e4-279">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-280">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-281">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-281">1.0</span></span>|
|[<span data-ttu-id="a85e4-282">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-283">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-284">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-285">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-285">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-286">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-286">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="a85e4-287">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="a85e4-287">(nullable) conversationId :String</span></span>

<span data-ttu-id="a85e4-288">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="a85e4-288">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="a85e4-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="a85e4-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-293">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-293">Type:</span></span>

*   <span data-ttu-id="a85e4-294">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a85e4-294">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-295">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-295">Requirements</span></span>

|<span data-ttu-id="a85e4-296">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-296">Requirement</span></span>|<span data-ttu-id="a85e4-297">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-298">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-299">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-299">1.0</span></span>|
|[<span data-ttu-id="a85e4-300">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-300">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-301">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-302">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-302">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-303">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-303">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="a85e4-304">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="a85e4-304">dateTimeCreated :Date</span></span>

<span data-ttu-id="a85e4-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-307">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-307">Type:</span></span>

*   <span data-ttu-id="a85e4-308">Data</span><span class="sxs-lookup"><span data-stu-id="a85e4-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-309">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-309">Requirements</span></span>

|<span data-ttu-id="a85e4-310">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-310">Requirement</span></span>|<span data-ttu-id="a85e4-311">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-312">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-313">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-313">1.0</span></span>|
|[<span data-ttu-id="a85e4-314">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-315">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-316">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-317">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-318">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-318">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="a85e4-319">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="a85e4-319">dateTimeModified :Date</span></span>

<span data-ttu-id="a85e4-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-322">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a85e4-322">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-323">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-323">Type:</span></span>

*   <span data-ttu-id="a85e4-324">Data</span><span class="sxs-lookup"><span data-stu-id="a85e4-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-325">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-325">Requirements</span></span>

|<span data-ttu-id="a85e4-326">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-326">Requirement</span></span>|<span data-ttu-id="a85e4-327">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-328">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-329">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-329">1.0</span></span>|
|[<span data-ttu-id="a85e4-330">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-331">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-332">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-333">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-334">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-334">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="a85e4-335">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a85e4-335">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="a85e4-336">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="a85e4-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="a85e4-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a85e4-339">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-339">Read mode</span></span>

<span data-ttu-id="a85e4-340">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-340">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a85e4-341">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a85e4-341">Compose mode</span></span>

<span data-ttu-id="a85e4-342">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="a85e4-343">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="a85e4-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-344">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-344">Type:</span></span>

*   <span data-ttu-id="a85e4-345">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a85e4-345">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-346">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-346">Requirements</span></span>

|<span data-ttu-id="a85e4-347">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-347">Requirement</span></span>|<span data-ttu-id="a85e4-348">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-348">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-349">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-349">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-350">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-350">1.0</span></span>|
|[<span data-ttu-id="a85e4-351">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-351">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-352">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-352">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-353">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-353">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-354">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-354">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-355">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-355">Example</span></span>

<span data-ttu-id="a85e4-356">O exemplo a seguir define a hora de término de um compromisso no modo de redação usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-356">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="a85e4-357">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="a85e4-357">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="a85e4-358">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a85e4-358">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="a85e4-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-361">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-361">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a85e4-362">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-362">Read mode</span></span>

<span data-ttu-id="a85e4-363">A propriedade `from` retorna um objeto `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-363">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="a85e4-364">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a85e4-364">Compose mode</span></span>

<span data-ttu-id="a85e4-365">A propriedade `from` retorna um objeto `From` que fornece um método para obtenção do valor de from.</span><span class="sxs-lookup"><span data-stu-id="a85e4-365">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a85e4-366">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-366">Type:</span></span>

*   <span data-ttu-id="a85e4-367">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="a85e4-367">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-368">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-368">Requirements</span></span>

|<span data-ttu-id="a85e4-369">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-369">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a85e4-370">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-371">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-371">1.0</span></span>|<span data-ttu-id="a85e4-372">1.7</span><span class="sxs-lookup"><span data-stu-id="a85e4-372">1.7</span></span>|
|[<span data-ttu-id="a85e4-373">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-373">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-374">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-374">ReadItem</span></span>|<span data-ttu-id="a85e4-375">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-375">ReadWriteItem</span></span>|
|[<span data-ttu-id="a85e4-376">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-376">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-377">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-377">Read</span></span>|<span data-ttu-id="a85e4-378">Escrever</span><span class="sxs-lookup"><span data-stu-id="a85e4-378">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="a85e4-379">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="a85e4-379">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="a85e4-380">Obtém ou define os cabeçalhos de internet de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a85e4-380">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-381">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-381">Type:</span></span>

*   [<span data-ttu-id="a85e4-382">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="a85e4-382">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="a85e4-383">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-383">Requirements</span></span>

|<span data-ttu-id="a85e4-384">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-384">Requirement</span></span>|<span data-ttu-id="a85e4-385">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-386">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-387">Visualização</span><span class="sxs-lookup"><span data-stu-id="a85e4-387">Preview</span></span>|
|[<span data-ttu-id="a85e4-388">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-389">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-390">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-391">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-391">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="a85e4-392">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="a85e4-392">internetMessageId :String</span></span>

<span data-ttu-id="a85e4-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-395">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-395">Type:</span></span>

*   <span data-ttu-id="a85e4-396">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a85e4-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-397">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-397">Requirements</span></span>

|<span data-ttu-id="a85e4-398">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-398">Requirement</span></span>|<span data-ttu-id="a85e4-399">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-400">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-401">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-401">1.0</span></span>|
|[<span data-ttu-id="a85e4-402">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-402">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-403">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-404">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-404">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-405">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-406">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-406">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="a85e4-407">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="a85e4-407">itemClass :String</span></span>

<span data-ttu-id="a85e4-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="a85e4-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="a85e4-412">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-412">Type</span></span>|<span data-ttu-id="a85e4-413">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-413">Description</span></span>|<span data-ttu-id="a85e4-414">classe de item</span><span class="sxs-lookup"><span data-stu-id="a85e4-414">item class</span></span>|
|---|---|---|
|<span data-ttu-id="a85e4-415">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="a85e4-415">Appointment items</span></span>|<span data-ttu-id="a85e4-416">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-416">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="a85e4-417">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="a85e4-417">Message items</span></span>|<span data-ttu-id="a85e4-418">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="a85e4-418">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="a85e4-419">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-419">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-420">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-420">Type:</span></span>

*   <span data-ttu-id="a85e4-421">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a85e4-421">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-422">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-422">Requirements</span></span>

|<span data-ttu-id="a85e4-423">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-423">Requirement</span></span>|<span data-ttu-id="a85e4-424">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-424">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-425">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-426">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-426">1.0</span></span>|
|[<span data-ttu-id="a85e4-427">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-428">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-429">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-430">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-430">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-431">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-431">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="a85e4-432">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="a85e4-432">(nullable) itemId :String</span></span>

<span data-ttu-id="a85e4-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-435">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="a85e4-435">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a85e4-436">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a85e4-436">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="a85e4-437">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a85e4-437">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a85e4-438">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="a85e4-438">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="a85e4-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-441">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-441">Type:</span></span>

*   <span data-ttu-id="a85e4-442">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a85e4-442">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-443">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-443">Requirements</span></span>

|<span data-ttu-id="a85e4-444">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-444">Requirement</span></span>|<span data-ttu-id="a85e4-445">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-446">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-447">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-447">1.0</span></span>|
|[<span data-ttu-id="a85e4-448">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-449">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-450">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-451">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-452">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-452">Example</span></span>

<span data-ttu-id="a85e4-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="a85e4-455">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="a85e4-455">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="a85e4-456">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="a85e4-456">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="a85e4-457">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a85e4-457">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-458">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-458">Type:</span></span>

*   [<span data-ttu-id="a85e4-459">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="a85e4-459">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="a85e4-460">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-460">Requirements</span></span>

|<span data-ttu-id="a85e4-461">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-461">Requirement</span></span>|<span data-ttu-id="a85e4-462">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-462">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-463">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-463">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-464">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-464">1.0</span></span>|
|[<span data-ttu-id="a85e4-465">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-465">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-466">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-466">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-467">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-467">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-468">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-468">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-469">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-469">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="a85e4-470">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="a85e4-470">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="a85e4-471">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a85e4-471">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a85e4-472">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-472">Read mode</span></span>

<span data-ttu-id="a85e4-473">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="a85e4-473">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a85e4-474">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a85e4-474">Compose mode</span></span>

<span data-ttu-id="a85e4-475">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="a85e4-475">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-476">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-476">Type:</span></span>

*   <span data-ttu-id="a85e4-477">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="a85e4-477">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-478">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-478">Requirements</span></span>

|<span data-ttu-id="a85e4-479">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-479">Requirement</span></span>|<span data-ttu-id="a85e4-480">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-481">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-482">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-482">1.0</span></span>|
|[<span data-ttu-id="a85e4-483">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-484">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-485">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-486">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-487">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-487">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="a85e4-488">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="a85e4-488">normalizedSubject :String</span></span>

<span data-ttu-id="a85e4-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="a85e4-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="a85e4-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-493">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-493">Type:</span></span>

*   <span data-ttu-id="a85e4-494">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a85e4-494">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-495">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-495">Requirements</span></span>

|<span data-ttu-id="a85e4-496">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-496">Requirement</span></span>|<span data-ttu-id="a85e4-497">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-497">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-498">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-498">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-499">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-499">1.0</span></span>|
|[<span data-ttu-id="a85e4-500">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-500">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-501">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-501">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-502">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-502">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-503">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-503">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-504">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-504">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="a85e4-505">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="a85e4-505">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="a85e4-506">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="a85e4-506">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-507">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-507">Type:</span></span>

*   [<span data-ttu-id="a85e4-508">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="a85e4-508">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="a85e4-509">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-509">Requirements</span></span>

|<span data-ttu-id="a85e4-510">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-510">Requirement</span></span>|<span data-ttu-id="a85e4-511">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-511">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-512">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-512">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-513">1.3</span><span class="sxs-lookup"><span data-stu-id="a85e4-513">1.3</span></span>|
|[<span data-ttu-id="a85e4-514">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-514">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-515">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-515">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-516">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-516">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-517">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-517">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a85e4-518">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a85e4-518">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a85e4-519">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="a85e4-519">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="a85e4-520">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a85e4-520">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a85e4-521">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-521">Read mode</span></span>

<span data-ttu-id="a85e4-522">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="a85e4-522">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a85e4-523">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a85e4-523">Compose mode</span></span>

<span data-ttu-id="a85e4-524">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="a85e4-524">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-525">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-525">Type:</span></span>

*   <span data-ttu-id="a85e4-526">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a85e4-526">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-527">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-527">Requirements</span></span>

|<span data-ttu-id="a85e4-528">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-528">Requirement</span></span>|<span data-ttu-id="a85e4-529">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-529">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-530">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-530">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-531">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-531">1.0</span></span>|
|[<span data-ttu-id="a85e4-532">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-532">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-533">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-533">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-534">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-534">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-535">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-535">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-536">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-536">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="a85e4-537">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="a85e4-537">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="a85e4-538">Obtém o endereço de email do organizador para uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-538">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a85e4-539">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-539">Read mode</span></span>

<span data-ttu-id="a85e4-540">A propriedade `organizer` retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="a85e4-540">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a85e4-541">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a85e4-541">Compose mode</span></span>

<span data-ttu-id="a85e4-542">A propriedade `organizer` retorna um objeto [Organizer](/javascript/api/outlook/office.organizer) que fornece um método para obtenção do valor de organizer.</span><span class="sxs-lookup"><span data-stu-id="a85e4-542">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-543">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-543">Type:</span></span>

*   <span data-ttu-id="a85e4-544">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="a85e4-544">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-545">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-545">Requirements</span></span>

|<span data-ttu-id="a85e4-546">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-546">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a85e4-547">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-547">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-548">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-548">1.0</span></span>|<span data-ttu-id="a85e4-549">1.7</span><span class="sxs-lookup"><span data-stu-id="a85e4-549">1.7</span></span>|
|[<span data-ttu-id="a85e4-550">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-550">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-551">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-551">ReadItem</span></span>|<span data-ttu-id="a85e4-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="a85e4-553">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-554">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-554">Read</span></span>|<span data-ttu-id="a85e4-555">Escrever</span><span class="sxs-lookup"><span data-stu-id="a85e4-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-556">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-556">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="a85e4-557">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="a85e4-557">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="a85e4-558">Obtém ou configura o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a85e4-558">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="a85e4-559">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="a85e4-559">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="a85e4-560">Modos de leitura e redação para itens do compromisso.</span><span class="sxs-lookup"><span data-stu-id="a85e4-560">Read and compose modes for appointment items.</span></span> <span data-ttu-id="a85e4-561">Modo de leitura para os itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="a85e4-561">Read mode for meeting request items.</span></span>

<span data-ttu-id="a85e4-562">A propriedade `recurrence` retorna um objeto [recurrence](/javascript/api/outlook/office.recurrence) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="a85e4-562">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="a85e4-563">`null` retorna para compromissos individuais e solicitações de reunião de compromissos individuais.</span><span class="sxs-lookup"><span data-stu-id="a85e4-563">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="a85e4-564">`undefined` retorna para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="a85e4-564">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="a85e4-565">Observação: solicitações de reunião têm um valor `itemClass` de IPM.Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="a85e4-565">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="a85e4-566">Observação: se o objeto de recorrência for `null`, isso indicará que o objeto é um compromisso individual ou uma solicitação de reunião de um compromisso individual e NÃO parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="a85e4-566">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-567">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-567">Type:</span></span>

* [<span data-ttu-id="a85e4-568">Recurrence</span><span class="sxs-lookup"><span data-stu-id="a85e4-568">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="a85e4-569">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-569">Requirement</span></span>|<span data-ttu-id="a85e4-570">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-570">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-571">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-571">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-572">1.7</span><span class="sxs-lookup"><span data-stu-id="a85e4-572">1.7</span></span>|
|[<span data-ttu-id="a85e4-573">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-573">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-574">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-574">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-575">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-575">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-576">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-576">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a85e4-577">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a85e4-577">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a85e4-578">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="a85e4-578">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="a85e4-579">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a85e4-579">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a85e4-580">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-580">Read mode</span></span>

<span data-ttu-id="a85e4-581">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="a85e4-581">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a85e4-582">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a85e4-582">Compose mode</span></span>

<span data-ttu-id="a85e4-583">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="a85e4-583">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-584">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-584">Type:</span></span>

*   <span data-ttu-id="a85e4-585">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a85e4-585">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-586">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-586">Requirements</span></span>

|<span data-ttu-id="a85e4-587">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-587">Requirement</span></span>|<span data-ttu-id="a85e4-588">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-588">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-589">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-589">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-590">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-590">1.0</span></span>|
|[<span data-ttu-id="a85e4-591">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-591">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-592">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-592">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-593">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-593">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-594">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-594">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-595">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-595">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="a85e4-596">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a85e4-596">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="a85e4-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="a85e4-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-601">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-601">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-602">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-602">Type:</span></span>

*   [<span data-ttu-id="a85e4-603">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a85e4-603">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a85e4-604">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-604">Requirements</span></span>

|<span data-ttu-id="a85e4-605">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-605">Requirement</span></span>|<span data-ttu-id="a85e4-606">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-607">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-608">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-608">1.0</span></span>|
|[<span data-ttu-id="a85e4-609">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-610">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-611">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-612">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-612">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-613">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-613">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="a85e4-614">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="a85e4-614">(nullable) seriesId :String</span></span>

<span data-ttu-id="a85e4-615">Obtém a id da série a qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="a85e4-615">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="a85e4-616">No OWA e no Outlook, o `seriesId` retorna a ID dos Serviços Web do Exchange (EWS) do item pai (série) a qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="a85e4-616">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="a85e4-617">No entanto, no iOS e no Android, o `seriesId` retorna a ID REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="a85e4-617">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-618">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="a85e4-618">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a85e4-619">A propriedade `seriesId` não é idêntica à ID do Outlook usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a85e4-619">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="a85e4-620">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a85e4-620">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a85e4-621">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="a85e4-621">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="a85e4-622">A propriedade `seriesId` retorna `null` para itens que não têm itens pai como compromissos individuais, itens de série ou solicitações de reunião e retorna `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="a85e4-622">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-623">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-623">Type:</span></span>

* <span data-ttu-id="a85e4-624">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a85e4-624">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-625">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-625">Requirements</span></span>

|<span data-ttu-id="a85e4-626">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-626">Requirement</span></span>|<span data-ttu-id="a85e4-627">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-628">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-628">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-629">1.7</span><span class="sxs-lookup"><span data-stu-id="a85e4-629">1.7</span></span>|
|[<span data-ttu-id="a85e4-630">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-631">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-632">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-633">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-633">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-634">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-634">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="a85e4-635">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a85e4-635">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="a85e4-636">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="a85e4-636">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="a85e4-p130">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a85e4-639">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-639">Read mode</span></span>

<span data-ttu-id="a85e4-640">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-640">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a85e4-641">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a85e4-641">Compose mode</span></span>

<span data-ttu-id="a85e4-642">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-642">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="a85e4-643">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="a85e4-643">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-644">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-644">Type:</span></span>

*   <span data-ttu-id="a85e4-645">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a85e4-645">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-646">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-646">Requirements</span></span>

|<span data-ttu-id="a85e4-647">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-647">Requirement</span></span>|<span data-ttu-id="a85e4-648">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-648">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-649">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-649">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-650">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-650">1.0</span></span>|
|[<span data-ttu-id="a85e4-651">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-651">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-652">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-652">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-653">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-653">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-654">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-654">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-655">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-655">Example</span></span>

<span data-ttu-id="a85e4-656">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-656">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="a85e4-657">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a85e4-657">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="a85e4-658">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="a85e4-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="a85e4-659">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="a85e4-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a85e4-660">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-660">Read mode</span></span>

<span data-ttu-id="a85e4-p131">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="a85e4-663">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a85e4-663">Compose mode</span></span>

<span data-ttu-id="a85e4-664">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="a85e4-664">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a85e4-665">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-665">Type:</span></span>

*   <span data-ttu-id="a85e4-666">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a85e4-666">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-667">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-667">Requirements</span></span>

|<span data-ttu-id="a85e4-668">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-668">Requirement</span></span>|<span data-ttu-id="a85e4-669">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-670">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-671">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-671">1.0</span></span>|
|[<span data-ttu-id="a85e4-672">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-672">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-673">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-674">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-674">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-675">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-675">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a85e4-676">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a85e4-676">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a85e4-677">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a85e4-677">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="a85e4-678">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a85e4-678">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a85e4-679">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-679">Read mode</span></span>

<span data-ttu-id="a85e4-p133">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a85e4-682">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a85e4-682">Compose mode</span></span>

<span data-ttu-id="a85e4-683">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="a85e4-683">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a85e4-684">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a85e4-684">Type:</span></span>

*   <span data-ttu-id="a85e4-685">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a85e4-685">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-686">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-686">Requirements</span></span>

|<span data-ttu-id="a85e4-687">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-687">Requirement</span></span>|<span data-ttu-id="a85e4-688">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-688">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-689">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-689">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-690">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-690">1.0</span></span>|
|[<span data-ttu-id="a85e4-691">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-691">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-692">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-692">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-693">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-693">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-694">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-694">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-695">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-695">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="a85e4-696">Métodos</span><span class="sxs-lookup"><span data-stu-id="a85e4-696">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="a85e4-697">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a85e4-697">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a85e4-698">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="a85e4-698">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a85e4-699">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="a85e4-699">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="a85e4-700">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a85e4-700">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-701">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-701">Parameters:</span></span>
|<span data-ttu-id="a85e4-702">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-702">Name</span></span>|<span data-ttu-id="a85e4-703">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-703">Type</span></span>|<span data-ttu-id="a85e4-704">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-704">Attributes</span></span>|<span data-ttu-id="a85e4-705">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-705">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="a85e4-706">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-706">String</span></span>||<span data-ttu-id="a85e4-p134">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a85e4-709">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-709">String</span></span>||<span data-ttu-id="a85e4-p135">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a85e4-712">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-712">Object</span></span>|<span data-ttu-id="a85e4-713">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-713">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-714">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-714">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a85e4-715">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-715">Object</span></span>|<span data-ttu-id="a85e4-716">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-716">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-717">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-717">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="a85e4-718">Booliano</span><span class="sxs-lookup"><span data-stu-id="a85e4-718">Boolean</span></span>|<span data-ttu-id="a85e4-719">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-719">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-720">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="a85e4-720">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="a85e4-721">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-721">function</span></span>|<span data-ttu-id="a85e4-722">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-722">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-723">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-723">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a85e4-724">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-724">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a85e4-725">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="a85e4-725">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a85e4-726">Erros</span><span class="sxs-lookup"><span data-stu-id="a85e4-726">Errors</span></span>

|<span data-ttu-id="a85e4-727">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a85e4-727">Error code</span></span>|<span data-ttu-id="a85e4-728">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-728">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="a85e4-729">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="a85e4-729">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="a85e4-730">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="a85e4-730">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a85e4-731">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="a85e4-731">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-732">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-732">Requirements</span></span>

|<span data-ttu-id="a85e4-733">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-733">Requirement</span></span>|<span data-ttu-id="a85e4-734">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-734">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-735">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-735">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-736">1.1</span><span class="sxs-lookup"><span data-stu-id="a85e4-736">1.1</span></span>|
|[<span data-ttu-id="a85e4-737">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-737">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-738">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-738">ReadWriteItem</span></span>|
|[<span data-ttu-id="a85e4-739">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-739">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-740">Escrever</span><span class="sxs-lookup"><span data-stu-id="a85e4-740">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a85e4-741">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a85e4-741">Examples</span></span>

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

<span data-ttu-id="a85e4-742">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="a85e4-742">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="a85e4-743">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a85e4-743">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a85e4-744">Adiciona um arquivo a partir da codificação base64 a uma mensagem ou compromisso como anexo.</span><span class="sxs-lookup"><span data-stu-id="a85e4-744">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a85e4-745">O método `addFileAttachmentFromBase64Async` carrega o arquivo a partir da codificação base64 e o anexa ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="a85e4-745">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="a85e4-746">Esse método retorna o identificador de anexo no objeto AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="a85e4-746">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="a85e4-747">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a85e4-747">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-748">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-748">Parameters:</span></span>
|<span data-ttu-id="a85e4-749">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-749">Name</span></span>|<span data-ttu-id="a85e4-750">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-750">Type</span></span>|<span data-ttu-id="a85e4-751">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-751">Attributes</span></span>|<span data-ttu-id="a85e4-752">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-752">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="a85e4-753">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-753">String</span></span>||<span data-ttu-id="a85e4-754">O conteúdo codificado em Base 64 de uma imagem ou arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="a85e4-754">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="a85e4-755">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-755">String</span></span>||<span data-ttu-id="a85e4-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a85e4-758">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-758">Object</span></span>|<span data-ttu-id="a85e4-759">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-759">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-760">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a85e4-761">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-761">Object</span></span>|<span data-ttu-id="a85e4-762">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-762">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-763">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="a85e4-764">Booliano</span><span class="sxs-lookup"><span data-stu-id="a85e4-764">Boolean</span></span>|<span data-ttu-id="a85e4-765">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-765">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-766">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="a85e4-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="a85e4-767">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-767">function</span></span>|<span data-ttu-id="a85e4-768">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-768">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-769">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a85e4-770">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a85e4-771">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="a85e4-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a85e4-772">Erros</span><span class="sxs-lookup"><span data-stu-id="a85e4-772">Errors</span></span>

|<span data-ttu-id="a85e4-773">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a85e4-773">Error code</span></span>|<span data-ttu-id="a85e4-774">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="a85e4-775">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="a85e4-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="a85e4-776">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="a85e4-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a85e4-777">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="a85e4-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-778">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-778">Requirements</span></span>

|<span data-ttu-id="a85e4-779">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-779">Requirement</span></span>|<span data-ttu-id="a85e4-780">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-781">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-782">Visualização</span><span class="sxs-lookup"><span data-stu-id="a85e4-782">Preview</span></span>|
|[<span data-ttu-id="a85e4-783">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-783">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="a85e4-785">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-785">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-786">Escrever</span><span class="sxs-lookup"><span data-stu-id="a85e4-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a85e4-787">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a85e4-787">Examples</span></span>

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
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a85e4-788">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a85e4-788">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a85e4-789">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="a85e4-789">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="a85e4-790">Atualmente, os tipos de evento compatíveis são `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-790">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-791">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-791">Parameters:</span></span>

| <span data-ttu-id="a85e4-792">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-792">Name</span></span> | <span data-ttu-id="a85e4-793">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-793">Type</span></span> | <span data-ttu-id="a85e4-794">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-794">Attributes</span></span> | <span data-ttu-id="a85e4-795">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-795">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a85e4-796">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a85e4-796">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a85e4-797">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="a85e4-797">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a85e4-798">Função</span><span class="sxs-lookup"><span data-stu-id="a85e4-798">Function</span></span> || <span data-ttu-id="a85e4-p138">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a85e4-802">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-802">Object</span></span> | <span data-ttu-id="a85e4-803">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-803">&lt;optional&gt;</span></span> | <span data-ttu-id="a85e4-804">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-804">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a85e4-805">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-805">Object</span></span> | <span data-ttu-id="a85e4-806">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-806">&lt;optional&gt;</span></span> | <span data-ttu-id="a85e4-807">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-807">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a85e4-808">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-808">function</span></span>| <span data-ttu-id="a85e4-809">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-809">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-810">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-810">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-811">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-811">Requirements</span></span>

|<span data-ttu-id="a85e4-812">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-812">Requirement</span></span>| <span data-ttu-id="a85e4-813">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-814">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a85e4-815">1.7</span><span class="sxs-lookup"><span data-stu-id="a85e4-815">1.7</span></span> |
|[<span data-ttu-id="a85e4-816">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a85e4-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-817">ReadItem</span></span> |
|[<span data-ttu-id="a85e4-818">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a85e4-819">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-819">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="a85e4-820">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a85e4-820">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a85e4-821">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="a85e4-821">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="a85e4-p139">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="a85e4-825">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a85e4-825">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="a85e4-826">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="a85e4-826">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-827">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-827">Parameters:</span></span>

|<span data-ttu-id="a85e4-828">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-828">Name</span></span>|<span data-ttu-id="a85e4-829">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-829">Type</span></span>|<span data-ttu-id="a85e4-830">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-830">Attributes</span></span>|<span data-ttu-id="a85e4-831">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-831">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="a85e4-832">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-832">String</span></span>||<span data-ttu-id="a85e4-p140">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a85e4-835">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-835">String</span></span>||<span data-ttu-id="a85e4-p141">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a85e4-838">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-838">Object</span></span>|<span data-ttu-id="a85e4-839">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-839">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-840">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-840">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a85e4-841">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-841">Object</span></span>|<span data-ttu-id="a85e4-842">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-842">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-843">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-843">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a85e4-844">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-844">function</span></span>|<span data-ttu-id="a85e4-845">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-845">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-846">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-846">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a85e4-847">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-847">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a85e4-848">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="a85e4-848">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a85e4-849">Erros</span><span class="sxs-lookup"><span data-stu-id="a85e4-849">Errors</span></span>

|<span data-ttu-id="a85e4-850">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a85e4-850">Error code</span></span>|<span data-ttu-id="a85e4-851">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-851">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a85e4-852">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="a85e4-852">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-853">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-853">Requirements</span></span>

|<span data-ttu-id="a85e4-854">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-854">Requirement</span></span>|<span data-ttu-id="a85e4-855">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-856">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-857">1.1</span><span class="sxs-lookup"><span data-stu-id="a85e4-857">1.1</span></span>|
|[<span data-ttu-id="a85e4-858">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-859">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-859">ReadWriteItem</span></span>|
|[<span data-ttu-id="a85e4-860">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-861">Escrever</span><span class="sxs-lookup"><span data-stu-id="a85e4-861">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-862">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-862">Example</span></span>

<span data-ttu-id="a85e4-863">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-863">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
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

####  <a name="close"></a><span data-ttu-id="a85e4-864">close()</span><span class="sxs-lookup"><span data-stu-id="a85e4-864">close()</span></span>

<span data-ttu-id="a85e4-865">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="a85e4-865">Closes the current item that is being composed.</span></span>

<span data-ttu-id="a85e4-p142">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-868">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="a85e4-868">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="a85e4-869">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="a85e4-869">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-870">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-870">Requirements</span></span>

|<span data-ttu-id="a85e4-871">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-871">Requirement</span></span>|<span data-ttu-id="a85e4-872">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-873">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-874">1.3</span><span class="sxs-lookup"><span data-stu-id="a85e4-874">1.3</span></span>|
|[<span data-ttu-id="a85e4-875">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-876">Restrito</span><span class="sxs-lookup"><span data-stu-id="a85e4-876">Restricted</span></span>|
|[<span data-ttu-id="a85e4-877">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-878">Escrever</span><span class="sxs-lookup"><span data-stu-id="a85e4-878">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="a85e4-879">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a85e4-879">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="a85e4-880">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="a85e4-880">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-881">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a85e4-881">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a85e4-882">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="a85e4-882">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a85e4-883">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="a85e4-883">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="a85e4-p143">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-887">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-887">Parameters:</span></span>

|<span data-ttu-id="a85e4-888">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-888">Name</span></span>|<span data-ttu-id="a85e4-889">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-889">Type</span></span>|<span data-ttu-id="a85e4-890">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-890">Attributes</span></span>|<span data-ttu-id="a85e4-891">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-891">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a85e4-892">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a85e4-892">String &#124; Object</span></span>||<span data-ttu-id="a85e4-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a85e4-895">**OU**</span><span class="sxs-lookup"><span data-stu-id="a85e4-895">**OR**</span></span><br/><span data-ttu-id="a85e4-p145">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a85e4-898">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-898">String</span></span>|<span data-ttu-id="a85e4-899">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-899">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a85e4-902">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-902">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a85e4-903">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-903">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-904">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="a85e4-904">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a85e4-905">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-905">String</span></span>||<span data-ttu-id="a85e4-p147">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a85e4-908">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-908">String</span></span>||<span data-ttu-id="a85e4-909">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="a85e4-909">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a85e4-910">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-910">String</span></span>||<span data-ttu-id="a85e4-p148">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a85e4-913">Booliano</span><span class="sxs-lookup"><span data-stu-id="a85e4-913">Boolean</span></span>||<span data-ttu-id="a85e4-p149">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a85e4-916">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-916">String</span></span>||<span data-ttu-id="a85e4-p150">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a85e4-920">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-920">function</span></span>|<span data-ttu-id="a85e4-921">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-921">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-922">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-922">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-923">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-923">Requirements</span></span>

|<span data-ttu-id="a85e4-924">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-924">Requirement</span></span>|<span data-ttu-id="a85e4-925">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-925">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-926">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-926">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-927">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-927">1.0</span></span>|
|[<span data-ttu-id="a85e4-928">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-928">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-929">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-929">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-930">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-930">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-931">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-931">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a85e4-932">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a85e4-932">Examples</span></span>

<span data-ttu-id="a85e4-933">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-933">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="a85e4-934">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="a85e4-934">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="a85e4-935">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="a85e4-935">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a85e4-936">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="a85e4-936">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a85e4-937">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="a85e4-937">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a85e4-938">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-938">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="a85e4-939">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a85e4-939">displayReplyForm(formData)</span></span>

<span data-ttu-id="a85e4-940">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="a85e4-940">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-941">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a85e4-941">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a85e4-942">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="a85e4-942">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a85e4-943">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="a85e4-943">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="a85e4-p151">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-947">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-947">Parameters:</span></span>

|<span data-ttu-id="a85e4-948">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-948">Name</span></span>|<span data-ttu-id="a85e4-949">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-949">Type</span></span>|<span data-ttu-id="a85e4-950">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-950">Attributes</span></span>|<span data-ttu-id="a85e4-951">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-951">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a85e4-952">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a85e4-952">String &#124; Object</span></span>||<span data-ttu-id="a85e4-p152">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a85e4-955">**OU**</span><span class="sxs-lookup"><span data-stu-id="a85e4-955">**OR**</span></span><br/><span data-ttu-id="a85e4-p153">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a85e4-958">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-958">String</span></span>|<span data-ttu-id="a85e4-959">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-959">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a85e4-962">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-962">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a85e4-963">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-963">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-964">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="a85e4-964">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a85e4-965">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-965">String</span></span>||<span data-ttu-id="a85e4-p155">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a85e4-968">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-968">String</span></span>||<span data-ttu-id="a85e4-969">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="a85e4-969">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a85e4-970">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-970">String</span></span>||<span data-ttu-id="a85e4-p156">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a85e4-973">Booliano</span><span class="sxs-lookup"><span data-stu-id="a85e4-973">Boolean</span></span>||<span data-ttu-id="a85e4-p157">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a85e4-976">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-976">String</span></span>||<span data-ttu-id="a85e4-p158">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a85e4-980">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-980">function</span></span>|<span data-ttu-id="a85e4-981">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-981">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-982">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-983">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-983">Requirements</span></span>

|<span data-ttu-id="a85e4-984">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-984">Requirement</span></span>|<span data-ttu-id="a85e4-985">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-986">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-987">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-987">1.0</span></span>|
|[<span data-ttu-id="a85e4-988">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-989">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-989">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-990">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-991">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-991">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a85e4-992">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a85e4-992">Examples</span></span>

<span data-ttu-id="a85e4-993">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-993">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="a85e4-994">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="a85e4-994">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="a85e4-995">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="a85e4-995">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a85e4-996">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="a85e4-996">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a85e4-997">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="a85e4-997">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a85e4-998">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-998">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="a85e4-999">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="a85e4-999">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="a85e4-1000">Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um objeto `AttachmentContent`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1000">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="a85e4-1001">O método `getAttachmentContentAsync` remove o obtém anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1001">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="a85e4-1002">Como melhor prática, você deve usar o identificador para recuperar um anexo na mesma sessão da qual attachmentIds foram recuperadas com o chamada `getAttachmentsAsync` ou `item.attachments`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1002">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="a85e4-1003">No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1003">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="a85e4-1004">Uma sessão é finalizada quando o usuário fecha o aplicativo, ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1004">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-1005">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1005">Parameters:</span></span>

|<span data-ttu-id="a85e4-1006">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-1006">Name</span></span>|<span data-ttu-id="a85e4-1007">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1007">Type</span></span>|<span data-ttu-id="a85e4-1008">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1008">Attributes</span></span>|<span data-ttu-id="a85e4-1009">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1009">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="a85e4-1010">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-1010">String</span></span>||<span data-ttu-id="a85e4-1011">O identificador do anexo que você quer obter.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1011">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="a85e4-1012">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1012">Object</span></span>|<span data-ttu-id="a85e4-1013">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1013">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1014">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1014">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a85e4-1015">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1015">Object</span></span>|<span data-ttu-id="a85e4-1016">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1016">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1017">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1017">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a85e4-1018">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-1018">function</span></span>|<span data-ttu-id="a85e4-1019">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1019">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1020">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-1020">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-1021">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1021">Requirements</span></span>

|<span data-ttu-id="a85e4-1022">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1022">Requirement</span></span>|<span data-ttu-id="a85e4-1023">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1024">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1025">Visualização</span><span class="sxs-lookup"><span data-stu-id="a85e4-1025">Preview</span></span>|
|[<span data-ttu-id="a85e4-1026">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1026">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1027">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-1028">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1028">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1029">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-1029">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a85e4-1030">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1030">Returns:</span></span>

<span data-ttu-id="a85e4-1031">Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="a85e4-1031">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="a85e4-1032">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1032">Example</span></span>

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
    // parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file
    if (result.format == Office.MailboxEnums.AttachmentContentFormat.Base64) {
        // handle file attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.Eml) {
        // handle item attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
        // handle .icalender attachment
    }
    else {
        // handle cloud attachment  
    }
}
```

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="a85e4-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a85e4-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="a85e4-1034">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1034">Gets the item's attachments as an array.</span></span> <span data-ttu-id="a85e4-1035">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1035">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-1036">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1036">Parameters:</span></span>

|<span data-ttu-id="a85e4-1037">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-1037">Name</span></span>|<span data-ttu-id="a85e4-1038">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1038">Type</span></span>|<span data-ttu-id="a85e4-1039">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1039">Attributes</span></span>|<span data-ttu-id="a85e4-1040">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1040">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a85e4-1041">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1041">Object</span></span>|<span data-ttu-id="a85e4-1042">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1043">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1043">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a85e4-1044">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1044">Object</span></span>|<span data-ttu-id="a85e4-1045">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1046">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1046">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a85e4-1047">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-1047">function</span></span>|<span data-ttu-id="a85e4-1048">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1048">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1049">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-1049">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-1050">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1050">Requirements</span></span>

|<span data-ttu-id="a85e4-1051">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1051">Requirement</span></span>|<span data-ttu-id="a85e4-1052">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1053">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1054">Visualização</span><span class="sxs-lookup"><span data-stu-id="a85e4-1054">Preview</span></span>|
|[<span data-ttu-id="a85e4-1055">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1056">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-1057">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1058">Escrever</span><span class="sxs-lookup"><span data-stu-id="a85e4-1058">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a85e4-1059">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1059">Returns:</span></span>

<span data-ttu-id="a85e4-1060">Tipo: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a85e4-1060">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="a85e4-1061">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1061">Example</span></span>

<span data-ttu-id="a85e4-1062">O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1062">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);  
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var _att = result.value [i];
            outputString += "<BR>" + i + ". Name: ";
            outputString += _att.name;
            outputString += "<BR>ID: " + _att.id;
            outputString += "<BR>contentType: " + _att.contentType;
            outputString += "<BR>size: " + _att.size;
            outputString += "<BR>attachmentType: " + _att.attachmentType;
            outputString += "<BR>isInline: " + _att.isInline;
        }
    }
}
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="a85e4-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a85e4-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="a85e4-1064">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1064">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-1065">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1065">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-1066">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1066">Requirements</span></span>

|<span data-ttu-id="a85e4-1067">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1067">Requirement</span></span>|<span data-ttu-id="a85e4-1068">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1069">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1070">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-1070">1.0</span></span>|
|[<span data-ttu-id="a85e4-1071">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1071">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1072">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-1073">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1073">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1074">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-1074">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a85e4-1075">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1075">Returns:</span></span>

<span data-ttu-id="a85e4-1076">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a85e4-1076">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a85e4-1077">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1077">Example</span></span>

<span data-ttu-id="a85e4-1078">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1078">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="a85e4-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a85e4-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a85e4-1080">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1080">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-1081">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1081">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-1082">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1082">Parameters:</span></span>

|<span data-ttu-id="a85e4-1083">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-1083">Name</span></span>|<span data-ttu-id="a85e4-1084">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1084">Type</span></span>|<span data-ttu-id="a85e4-1085">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1085">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="a85e4-1086">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="a85e4-1086">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="a85e4-1087">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1087">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-1088">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1088">Requirements</span></span>

|<span data-ttu-id="a85e4-1089">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1089">Requirement</span></span>|<span data-ttu-id="a85e4-1090">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1091">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1092">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-1092">1.0</span></span>|
|[<span data-ttu-id="a85e4-1093">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1094">Restrito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1094">Restricted</span></span>|
|[<span data-ttu-id="a85e4-1095">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1096">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-1096">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a85e4-1097">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1097">Returns:</span></span>

<span data-ttu-id="a85e4-1098">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1098">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="a85e4-1099">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1099">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="a85e4-1100">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1100">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="a85e4-1101">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1101">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="a85e4-1102">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="a85e4-1102">Value of `entityType`</span></span>|<span data-ttu-id="a85e4-1103">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="a85e4-1103">Type of objects in returned array</span></span>|<span data-ttu-id="a85e4-1104">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="a85e4-1104">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="a85e4-1105">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-1105">String</span></span>|<span data-ttu-id="a85e4-1106">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="a85e4-1106">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="a85e4-1107">Contato</span><span class="sxs-lookup"><span data-stu-id="a85e4-1107">Contact</span></span>|<span data-ttu-id="a85e4-1108">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a85e4-1108">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="a85e4-1109">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-1109">String</span></span>|<span data-ttu-id="a85e4-1110">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a85e4-1110">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="a85e4-1111">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="a85e4-1111">MeetingSuggestion</span></span>|<span data-ttu-id="a85e4-1112">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a85e4-1112">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="a85e4-1113">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="a85e4-1113">PhoneNumber</span></span>|<span data-ttu-id="a85e4-1114">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="a85e4-1114">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="a85e4-1115">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="a85e4-1115">TaskSuggestion</span></span>|<span data-ttu-id="a85e4-1116">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a85e4-1116">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="a85e4-1117">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-1117">String</span></span>|<span data-ttu-id="a85e4-1118">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="a85e4-1118">**Restricted**</span></span>|

<span data-ttu-id="a85e4-1119">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a85e4-1119">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="a85e4-1120">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1120">Example</span></span>

<span data-ttu-id="a85e4-1121">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1121">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="a85e4-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a85e4-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a85e4-1123">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1123">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-1124">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1124">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a85e4-1125">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1125">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-1126">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1126">Parameters:</span></span>

|<span data-ttu-id="a85e4-1127">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-1127">Name</span></span>|<span data-ttu-id="a85e4-1128">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1128">Type</span></span>|<span data-ttu-id="a85e4-1129">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1129">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a85e4-1130">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-1130">String</span></span>|<span data-ttu-id="a85e4-1131">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1131">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-1132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1132">Requirements</span></span>

|<span data-ttu-id="a85e4-1133">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1133">Requirement</span></span>|<span data-ttu-id="a85e4-1134">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1136">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-1136">1.0</span></span>|
|[<span data-ttu-id="a85e4-1137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1138">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-1139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1140">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-1140">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a85e4-1141">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1141">Returns:</span></span>

<span data-ttu-id="a85e4-p162">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="a85e4-1144">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a85e4-1144">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="a85e4-1145">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a85e4-1145">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="a85e4-1146">Obtém dados de inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="a85e4-1146">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-1147">Esse método só é compatível com o Outlook 2016 ou posterior para Windows (versões Clique para Executar posteriores à 16.0.8413.1000) e o Outlook na Web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1147">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-1148">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1148">Parameters:</span></span>
|<span data-ttu-id="a85e4-1149">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-1149">Name</span></span>|<span data-ttu-id="a85e4-1150">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1150">Type</span></span>|<span data-ttu-id="a85e4-1151">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1151">Attributes</span></span>|<span data-ttu-id="a85e4-1152">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1152">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a85e4-1153">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1153">Object</span></span>|<span data-ttu-id="a85e4-1154">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1155">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1155">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a85e4-1156">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1156">Object</span></span>|<span data-ttu-id="a85e4-1157">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1158">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1158">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a85e4-1159">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-1159">function</span></span>|<span data-ttu-id="a85e4-1160">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1161">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-1161">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a85e4-1162">Após o êxito, os dados de inicialização são fornecidos na propriedade `asyncResult.value` como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1162">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="a85e4-1163">Se não houver nenhum contexto de inicialização, o objeto `asyncResult` conterá um objeto `Error` com sua propriedade `code` definida como `9020` e sua propriedade `name` definida como `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1163">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-1164">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1164">Requirements</span></span>

|<span data-ttu-id="a85e4-1165">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1165">Requirement</span></span>|<span data-ttu-id="a85e4-1166">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1166">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1167">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1167">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1168">Visualização</span><span class="sxs-lookup"><span data-stu-id="a85e4-1168">Preview</span></span>|
|[<span data-ttu-id="a85e4-1169">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1169">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1170">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1170">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-1171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1172">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-1172">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-1173">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1173">Example</span></span>

```javascript
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="a85e4-1174">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a85e4-1174">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="a85e4-1175">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1175">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-1176">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1176">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a85e4-p163">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a85e4-1180">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1180">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a85e4-1181">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1181">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a85e4-p164">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-1185">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1185">Requirements</span></span>

|<span data-ttu-id="a85e4-1186">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1186">Requirement</span></span>|<span data-ttu-id="a85e4-1187">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1188">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1189">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-1189">1.0</span></span>|
|[<span data-ttu-id="a85e4-1190">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1191">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-1192">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1193">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-1193">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a85e4-1194">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1194">Returns:</span></span>

<span data-ttu-id="a85e4-p165">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="a85e4-1197">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="a85e4-1197">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a85e4-1198">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1198">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a85e4-1199">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1199">Example</span></span>

<span data-ttu-id="a85e4-1200">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1200">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="a85e4-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="a85e4-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="a85e4-1202">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1202">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-1203">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1203">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a85e4-1204">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1204">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="a85e4-p166">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-1207">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1207">Parameters:</span></span>

|<span data-ttu-id="a85e4-1208">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-1208">Name</span></span>|<span data-ttu-id="a85e4-1209">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1209">Type</span></span>|<span data-ttu-id="a85e4-1210">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1210">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a85e4-1211">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-1211">String</span></span>|<span data-ttu-id="a85e4-1212">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1212">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-1213">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1213">Requirements</span></span>

|<span data-ttu-id="a85e4-1214">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1214">Requirement</span></span>|<span data-ttu-id="a85e4-1215">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1215">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1216">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1217">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-1217">1.0</span></span>|
|[<span data-ttu-id="a85e4-1218">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1219">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-1220">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1221">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-1221">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a85e4-1222">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1222">Returns:</span></span>

<span data-ttu-id="a85e4-1223">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1223">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="a85e4-1224">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="a85e4-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a85e4-1225">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="a85e4-1225">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a85e4-1226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1226">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="a85e4-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="a85e4-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="a85e4-1228">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1228">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="a85e4-p167">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-1231">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1231">Parameters:</span></span>

|<span data-ttu-id="a85e4-1232">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-1232">Name</span></span>|<span data-ttu-id="a85e4-1233">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1233">Type</span></span>|<span data-ttu-id="a85e4-1234">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1234">Attributes</span></span>|<span data-ttu-id="a85e4-1235">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1235">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="a85e4-1236">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a85e4-1236">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="a85e4-p168">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="a85e4-1240">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1240">Object</span></span>|<span data-ttu-id="a85e4-1241">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1241">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1242">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1242">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a85e4-1243">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1243">Object</span></span>|<span data-ttu-id="a85e4-1244">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1244">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1245">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1245">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a85e4-1246">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-1246">function</span></span>||<span data-ttu-id="a85e4-1247">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-1247">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a85e4-1248">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1248">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="a85e4-1249">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1249">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-1250">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1250">Requirements</span></span>

|<span data-ttu-id="a85e4-1251">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1251">Requirement</span></span>|<span data-ttu-id="a85e4-1252">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1252">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1253">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1254">1.2</span><span class="sxs-lookup"><span data-stu-id="a85e4-1254">1.2</span></span>|
|[<span data-ttu-id="a85e4-1255">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1256">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1256">ReadWriteItem</span></span>|
|[<span data-ttu-id="a85e4-1257">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1258">Escrever</span><span class="sxs-lookup"><span data-stu-id="a85e4-1258">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a85e4-1259">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1259">Returns:</span></span>

<span data-ttu-id="a85e4-1260">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1260">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="a85e4-1261">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="a85e4-1261">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a85e4-1262">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-1262">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a85e4-1263">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1263">Example</span></span>

```javascript
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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="a85e4-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a85e4-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="a85e4-p170">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a85e4-p170">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-1267">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1267">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-1268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1268">Requirements</span></span>

|<span data-ttu-id="a85e4-1269">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1269">Requirement</span></span>|<span data-ttu-id="a85e4-1270">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1271">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1272">1.6</span><span class="sxs-lookup"><span data-stu-id="a85e4-1272">1.6</span></span>|
|[<span data-ttu-id="a85e4-1273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1274">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-1275">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1276">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a85e4-1277">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1277">Returns:</span></span>

<span data-ttu-id="a85e4-1278">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a85e4-1278">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a85e4-1279">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1279">Example</span></span>

<span data-ttu-id="a85e4-1280">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1280">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="a85e4-1281">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a85e4-1281">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="a85e4-p171">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a85e4-p171">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-1284">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1284">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a85e4-p172">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p172">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a85e4-1288">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1288">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a85e4-1289">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1289">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a85e4-p173">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p173">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a85e4-1293">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1293">Requirements</span></span>

|<span data-ttu-id="a85e4-1294">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1294">Requirement</span></span>|<span data-ttu-id="a85e4-1295">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1295">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1296">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1296">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1297">1.6</span><span class="sxs-lookup"><span data-stu-id="a85e4-1297">1.6</span></span>|
|[<span data-ttu-id="a85e4-1298">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1298">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1299">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1299">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-1300">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1300">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1301">Read</span><span class="sxs-lookup"><span data-stu-id="a85e4-1301">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a85e4-1302">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1302">Returns:</span></span>

<span data-ttu-id="a85e4-p174">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p174">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="a85e4-1305">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1305">Example</span></span>

<span data-ttu-id="a85e4-1306">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1306">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="a85e4-1307">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="a85e4-1307">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="a85e4-1308">Obtém as propriedades do compromisso ou mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1308">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-1309">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1309">Parameters:</span></span>

|<span data-ttu-id="a85e4-1310">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-1310">Name</span></span>|<span data-ttu-id="a85e4-1311">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1311">Type</span></span>|<span data-ttu-id="a85e4-1312">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1312">Attributes</span></span>|<span data-ttu-id="a85e4-1313">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1313">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a85e4-1314">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1314">Object</span></span>|<span data-ttu-id="a85e4-1315">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1315">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1316">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1316">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a85e4-1317">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1317">Object</span></span>|<span data-ttu-id="a85e4-1318">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1318">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1319">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1319">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a85e4-1320">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-1320">function</span></span>||<span data-ttu-id="a85e4-1321">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-1321">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a85e4-1322">As propriedades compartilhadas são fornecidas como um objeto [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1322">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a85e4-1323">Esse objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1323">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-1324">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1324">Requirements</span></span>

|<span data-ttu-id="a85e4-1325">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1325">Requirement</span></span>|<span data-ttu-id="a85e4-1326">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1326">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1327">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1328">Visualização</span><span class="sxs-lookup"><span data-stu-id="a85e4-1328">Preview</span></span>|
|[<span data-ttu-id="a85e4-1329">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1330">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-1331">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1332">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-1332">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-1333">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1333">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="a85e4-1334">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a85e4-1334">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="a85e4-1335">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1335">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="a85e4-p176">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-1339">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1339">Parameters:</span></span>

|<span data-ttu-id="a85e4-1340">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-1340">Name</span></span>|<span data-ttu-id="a85e4-1341">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1341">Type</span></span>|<span data-ttu-id="a85e4-1342">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1342">Attributes</span></span>|<span data-ttu-id="a85e4-1343">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1343">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="a85e4-1344">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-1344">function</span></span>||<span data-ttu-id="a85e4-1345">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-1345">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a85e4-1346">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1346">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a85e4-1347">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1347">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="a85e4-1348">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1348">Object</span></span>|<span data-ttu-id="a85e4-1349">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1349">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1350">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1350">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="a85e4-1351">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1351">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-1352">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1352">Requirements</span></span>

|<span data-ttu-id="a85e4-1353">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1353">Requirement</span></span>|<span data-ttu-id="a85e4-1354">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1354">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1355">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1356">1.0</span><span class="sxs-lookup"><span data-stu-id="a85e4-1356">1.0</span></span>|
|[<span data-ttu-id="a85e4-1357">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1358">ReadItem</span></span>|
|[<span data-ttu-id="a85e4-1359">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1360">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-1360">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-1361">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1361">Example</span></span>

<span data-ttu-id="a85e4-p179">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="a85e4-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a85e4-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="a85e4-1366">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1366">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="a85e4-1367">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1367">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="a85e4-1368">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1368">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="a85e4-1369">No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1369">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="a85e4-1370">Uma sessão é finalizada quando o usuário fecha o aplicativo, ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1370">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-1371">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1371">Parameters:</span></span>

|<span data-ttu-id="a85e4-1372">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-1372">Name</span></span>|<span data-ttu-id="a85e4-1373">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1373">Type</span></span>|<span data-ttu-id="a85e4-1374">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1374">Attributes</span></span>|<span data-ttu-id="a85e4-1375">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1375">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="a85e4-1376">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-1376">String</span></span>||<span data-ttu-id="a85e4-1377">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1377">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="a85e4-1378">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1378">Object</span></span>|<span data-ttu-id="a85e4-1379">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1379">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1380">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1380">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a85e4-1381">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1381">Object</span></span>|<span data-ttu-id="a85e4-1382">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1382">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1383">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1383">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a85e4-1384">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-1384">function</span></span>|<span data-ttu-id="a85e4-1385">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1385">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1386">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-1386">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a85e4-1387">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1387">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a85e4-1388">Erros</span><span class="sxs-lookup"><span data-stu-id="a85e4-1388">Errors</span></span>

|<span data-ttu-id="a85e4-1389">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a85e4-1389">Error code</span></span>|<span data-ttu-id="a85e4-1390">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1390">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="a85e4-1391">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1391">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-1392">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1392">Requirements</span></span>

|<span data-ttu-id="a85e4-1393">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1393">Requirement</span></span>|<span data-ttu-id="a85e4-1394">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1394">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1395">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1396">1.1</span><span class="sxs-lookup"><span data-stu-id="a85e4-1396">1.1</span></span>|
|[<span data-ttu-id="a85e4-1397">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1398">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1398">ReadWriteItem</span></span>|
|[<span data-ttu-id="a85e4-1399">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1400">Escrever</span><span class="sxs-lookup"><span data-stu-id="a85e4-1400">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-1401">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1401">Example</span></span>

<span data-ttu-id="a85e4-1402">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1402">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="a85e4-1403">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a85e4-1403">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="a85e4-1404">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1404">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="a85e4-1405">Atualmente, os tipos de evento compatíveis são `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1405">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-1406">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1406">Parameters:</span></span>

| <span data-ttu-id="a85e4-1407">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-1407">Name</span></span> | <span data-ttu-id="a85e4-1408">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1408">Type</span></span> | <span data-ttu-id="a85e4-1409">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1409">Attributes</span></span> | <span data-ttu-id="a85e4-1410">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1410">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a85e4-1411">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a85e4-1411">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a85e4-1412">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1412">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="a85e4-1413">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1413">Object</span></span> | <span data-ttu-id="a85e4-1414">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1414">&lt;optional&gt;</span></span> | <span data-ttu-id="a85e4-1415">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1415">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a85e4-1416">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1416">Object</span></span> | <span data-ttu-id="a85e4-1417">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1417">&lt;optional&gt;</span></span> | <span data-ttu-id="a85e4-1418">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1418">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a85e4-1419">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-1419">function</span></span>| <span data-ttu-id="a85e4-1420">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1420">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1421">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-1421">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-1422">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1422">Requirements</span></span>

|<span data-ttu-id="a85e4-1423">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1423">Requirement</span></span>| <span data-ttu-id="a85e4-1424">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1424">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1425">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a85e4-1426">1.7</span><span class="sxs-lookup"><span data-stu-id="a85e4-1426">1.7</span></span> |
|[<span data-ttu-id="a85e4-1427">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a85e4-1428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1428">ReadItem</span></span> |
|[<span data-ttu-id="a85e4-1429">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a85e4-1430">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a85e4-1430">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="a85e4-1431">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="a85e4-1431">saveAsync([options], callback)</span></span>

<span data-ttu-id="a85e4-1432">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1432">Asynchronously saves an item.</span></span>

<span data-ttu-id="a85e4-p181">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-1436">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1436">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="a85e4-1437">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1437">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="a85e4-p183">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="a85e4-1441">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1441">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="a85e4-1442">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1442">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="a85e4-1443">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1443">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="a85e4-1444">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1444">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-1445">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1445">Parameters:</span></span>

|<span data-ttu-id="a85e4-1446">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-1446">Name</span></span>|<span data-ttu-id="a85e4-1447">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1447">Type</span></span>|<span data-ttu-id="a85e4-1448">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1448">Attributes</span></span>|<span data-ttu-id="a85e4-1449">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1449">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a85e4-1450">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1450">Object</span></span>|<span data-ttu-id="a85e4-1451">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1451">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1452">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1452">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a85e4-1453">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1453">Object</span></span>|<span data-ttu-id="a85e4-1454">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1454">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1455">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1455">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a85e4-1456">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-1456">function</span></span>||<span data-ttu-id="a85e4-1457">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-1457">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a85e4-1458">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1458">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-1459">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1459">Requirements</span></span>

|<span data-ttu-id="a85e4-1460">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1460">Requirement</span></span>|<span data-ttu-id="a85e4-1461">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1461">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1462">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1463">1.3</span><span class="sxs-lookup"><span data-stu-id="a85e4-1463">1.3</span></span>|
|[<span data-ttu-id="a85e4-1464">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1464">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1465">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1465">ReadWriteItem</span></span>|
|[<span data-ttu-id="a85e4-1466">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1466">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1467">Escrever</span><span class="sxs-lookup"><span data-stu-id="a85e4-1467">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a85e4-1468">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1468">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="a85e4-p185">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="a85e4-1471">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="a85e4-1471">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="a85e4-1472">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1472">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="a85e4-p186">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a85e4-1476">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a85e4-1476">Parameters:</span></span>

|<span data-ttu-id="a85e4-1477">Nome</span><span class="sxs-lookup"><span data-stu-id="a85e4-1477">Name</span></span>|<span data-ttu-id="a85e4-1478">Tipo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1478">Type</span></span>|<span data-ttu-id="a85e4-1479">Atributos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1479">Attributes</span></span>|<span data-ttu-id="a85e4-1480">Descrição</span><span class="sxs-lookup"><span data-stu-id="a85e4-1480">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="a85e4-1481">String</span><span class="sxs-lookup"><span data-stu-id="a85e4-1481">String</span></span>||<span data-ttu-id="a85e4-p187">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="a85e4-1485">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1485">Object</span></span>|<span data-ttu-id="a85e4-1486">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1486">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1487">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1487">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a85e4-1488">Objeto</span><span class="sxs-lookup"><span data-stu-id="a85e4-1488">Object</span></span>|<span data-ttu-id="a85e4-1489">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1489">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-1490">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1490">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="a85e4-1491">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a85e4-1491">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="a85e4-1492">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a85e4-1492">&lt;optional&gt;</span></span>|<span data-ttu-id="a85e4-p188">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="a85e4-p189">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="a85e4-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="a85e4-1497">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="a85e4-1497">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="a85e4-1498">function</span><span class="sxs-lookup"><span data-stu-id="a85e4-1498">function</span></span>||<span data-ttu-id="a85e4-1499">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a85e4-1499">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a85e4-1500">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a85e4-1500">Requirements</span></span>

|<span data-ttu-id="a85e4-1501">Requisito</span><span class="sxs-lookup"><span data-stu-id="a85e4-1501">Requirement</span></span>|<span data-ttu-id="a85e4-1502">Valor</span><span class="sxs-lookup"><span data-stu-id="a85e4-1502">Value</span></span>|
|---|---|
|[<span data-ttu-id="a85e4-1503">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a85e4-1503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a85e4-1504">1.2</span><span class="sxs-lookup"><span data-stu-id="a85e4-1504">1.2</span></span>|
|[<span data-ttu-id="a85e4-1505">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a85e4-1506">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a85e4-1506">ReadWriteItem</span></span>|
|[<span data-ttu-id="a85e4-1507">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a85e4-1507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a85e4-1508">Escrever</span><span class="sxs-lookup"><span data-stu-id="a85e4-1508">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a85e4-1509">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a85e4-1509">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
