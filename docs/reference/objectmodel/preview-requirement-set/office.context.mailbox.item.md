---
title: Office.Context.Mailbox.item - conjunto de requisições de visualização
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: a660f8bafdd2587f97d704e42c47abbe6c7d533d
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/13/2019
ms.locfileid: "29982045"
---
# <a name="item"></a><span data-ttu-id="af492-102">item</span><span class="sxs-lookup"><span data-stu-id="af492-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="af492-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="af492-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="af492-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="af492-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-106">Requirements</span></span>

|<span data-ttu-id="af492-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-107">Requirement</span></span>|<span data-ttu-id="af492-108">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-110">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-110">1.0</span></span>|
|[<span data-ttu-id="af492-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="af492-112">Restricted</span></span>|
|[<span data-ttu-id="af492-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-114">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="af492-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="af492-115">Members and methods</span></span>

| <span data-ttu-id="af492-116">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-116">Member</span></span> | <span data-ttu-id="af492-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="af492-118">attachments</span><span class="sxs-lookup"><span data-stu-id="af492-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="af492-119">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-119">Member</span></span> |
| [<span data-ttu-id="af492-120">bcc</span><span class="sxs-lookup"><span data-stu-id="af492-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="af492-121">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-121">Member</span></span> |
| [<span data-ttu-id="af492-122">body</span><span class="sxs-lookup"><span data-stu-id="af492-122">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="af492-123">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-123">Member</span></span> |
| [<span data-ttu-id="af492-124">cc</span><span class="sxs-lookup"><span data-stu-id="af492-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="af492-125">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-125">Member</span></span> |
| [<span data-ttu-id="af492-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="af492-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="af492-127">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-127">Member</span></span> |
| [<span data-ttu-id="af492-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="af492-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="af492-129">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-129">Member</span></span> |
| [<span data-ttu-id="af492-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="af492-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="af492-131">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-131">Member</span></span> |
| [<span data-ttu-id="af492-132">end</span><span class="sxs-lookup"><span data-stu-id="af492-132">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="af492-133">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-133">Member</span></span> |
| [<span data-ttu-id="af492-134">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="af492-134">enhancedLocation</span></span>](#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation) | <span data-ttu-id="af492-135">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-135">Member</span></span> |
| [<span data-ttu-id="af492-136">from</span><span class="sxs-lookup"><span data-stu-id="af492-136">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="af492-137">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-137">Member</span></span> |
| [<span data-ttu-id="af492-138">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="af492-138">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="af492-139">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-139">Member</span></span> |
| [<span data-ttu-id="af492-140">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="af492-140">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="af492-141">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-141">Member</span></span> |
| [<span data-ttu-id="af492-142">itemClass</span><span class="sxs-lookup"><span data-stu-id="af492-142">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="af492-143">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-143">Member</span></span> |
| [<span data-ttu-id="af492-144">itemId</span><span class="sxs-lookup"><span data-stu-id="af492-144">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="af492-145">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-145">Member</span></span> |
| [<span data-ttu-id="af492-146">itemType</span><span class="sxs-lookup"><span data-stu-id="af492-146">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="af492-147">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-147">Member</span></span> |
| [<span data-ttu-id="af492-148">location</span><span class="sxs-lookup"><span data-stu-id="af492-148">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="af492-149">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-149">Member</span></span> |
| [<span data-ttu-id="af492-150">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="af492-150">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="af492-151">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-151">Member</span></span> |
| [<span data-ttu-id="af492-152">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="af492-152">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="af492-153">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-153">Member</span></span> |
| [<span data-ttu-id="af492-154">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="af492-154">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="af492-155">Member</span><span class="sxs-lookup"><span data-stu-id="af492-155">Member</span></span> |
| [<span data-ttu-id="af492-156">organizer</span><span class="sxs-lookup"><span data-stu-id="af492-156">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="af492-157">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-157">Member</span></span> |
| [<span data-ttu-id="af492-158">recurrence</span><span class="sxs-lookup"><span data-stu-id="af492-158">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="af492-159">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-159">Member</span></span> |
| [<span data-ttu-id="af492-160">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="af492-160">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="af492-161">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-161">Member</span></span> |
| [<span data-ttu-id="af492-162">sender</span><span class="sxs-lookup"><span data-stu-id="af492-162">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="af492-163">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-163">Member</span></span> |
| [<span data-ttu-id="af492-164">seriesId</span><span class="sxs-lookup"><span data-stu-id="af492-164">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="af492-165">Member</span><span class="sxs-lookup"><span data-stu-id="af492-165">Member</span></span> |
| [<span data-ttu-id="af492-166">start</span><span class="sxs-lookup"><span data-stu-id="af492-166">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="af492-167">Member</span><span class="sxs-lookup"><span data-stu-id="af492-167">Member</span></span> |
| [<span data-ttu-id="af492-168">subject</span><span class="sxs-lookup"><span data-stu-id="af492-168">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="af492-169">Member</span><span class="sxs-lookup"><span data-stu-id="af492-169">Member</span></span> |
| [<span data-ttu-id="af492-170">to</span><span class="sxs-lookup"><span data-stu-id="af492-170">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="af492-171">Membro</span><span class="sxs-lookup"><span data-stu-id="af492-171">Member</span></span> |
| [<span data-ttu-id="af492-172">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="af492-172">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="af492-173">Método</span><span class="sxs-lookup"><span data-stu-id="af492-173">Method</span></span> |
| [<span data-ttu-id="af492-174">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="af492-174">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="af492-175">Método</span><span class="sxs-lookup"><span data-stu-id="af492-175">Method</span></span> |
| [<span data-ttu-id="af492-176">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="af492-176">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="af492-177">Método</span><span class="sxs-lookup"><span data-stu-id="af492-177">Method</span></span> |
| [<span data-ttu-id="af492-178">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="af492-178">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="af492-179">Método</span><span class="sxs-lookup"><span data-stu-id="af492-179">Method</span></span> |
| [<span data-ttu-id="af492-180">close</span><span class="sxs-lookup"><span data-stu-id="af492-180">close</span></span>](#close) | <span data-ttu-id="af492-181">Método</span><span class="sxs-lookup"><span data-stu-id="af492-181">Method</span></span> |
| [<span data-ttu-id="af492-182">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="af492-182">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="af492-183">Método</span><span class="sxs-lookup"><span data-stu-id="af492-183">Method</span></span> |
| [<span data-ttu-id="af492-184">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="af492-184">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="af492-185">Método</span><span class="sxs-lookup"><span data-stu-id="af492-185">Method</span></span> |
| [<span data-ttu-id="af492-186">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="af492-186">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="af492-187">Método</span><span class="sxs-lookup"><span data-stu-id="af492-187">Method</span></span> |
| [<span data-ttu-id="af492-188">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="af492-188">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="af492-189">Método</span><span class="sxs-lookup"><span data-stu-id="af492-189">Method</span></span> |
| [<span data-ttu-id="af492-190">getEntities</span><span class="sxs-lookup"><span data-stu-id="af492-190">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="af492-191">Método</span><span class="sxs-lookup"><span data-stu-id="af492-191">Method</span></span> |
| [<span data-ttu-id="af492-192">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="af492-192">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="af492-193">Método</span><span class="sxs-lookup"><span data-stu-id="af492-193">Method</span></span> |
| [<span data-ttu-id="af492-194">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="af492-194">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="af492-195">Método</span><span class="sxs-lookup"><span data-stu-id="af492-195">Method</span></span> |
| [<span data-ttu-id="af492-196">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="af492-196">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="af492-197">Método</span><span class="sxs-lookup"><span data-stu-id="af492-197">Method</span></span> |
| [<span data-ttu-id="af492-198">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="af492-198">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="af492-199">Método</span><span class="sxs-lookup"><span data-stu-id="af492-199">Method</span></span> |
| [<span data-ttu-id="af492-200">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="af492-200">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="af492-201">Método</span><span class="sxs-lookup"><span data-stu-id="af492-201">Method</span></span> |
| [<span data-ttu-id="af492-202">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="af492-202">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="af492-203">Método</span><span class="sxs-lookup"><span data-stu-id="af492-203">Method</span></span> |
| [<span data-ttu-id="af492-204">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="af492-204">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="af492-205">Método</span><span class="sxs-lookup"><span data-stu-id="af492-205">Method</span></span> |
| [<span data-ttu-id="af492-206">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="af492-206">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="af492-207">Método</span><span class="sxs-lookup"><span data-stu-id="af492-207">Method</span></span> |
| [<span data-ttu-id="af492-208">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="af492-208">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="af492-209">Método</span><span class="sxs-lookup"><span data-stu-id="af492-209">Method</span></span> |
| [<span data-ttu-id="af492-210">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="af492-210">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="af492-211">Método</span><span class="sxs-lookup"><span data-stu-id="af492-211">Method</span></span> |
| [<span data-ttu-id="af492-212">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="af492-212">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="af492-213">Método</span><span class="sxs-lookup"><span data-stu-id="af492-213">Method</span></span> |
| [<span data-ttu-id="af492-214">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="af492-214">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="af492-215">Método</span><span class="sxs-lookup"><span data-stu-id="af492-215">Method</span></span> |
| [<span data-ttu-id="af492-216">saveAsync</span><span class="sxs-lookup"><span data-stu-id="af492-216">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="af492-217">Método</span><span class="sxs-lookup"><span data-stu-id="af492-217">Method</span></span> |
| [<span data-ttu-id="af492-218">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="af492-218">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="af492-219">Método</span><span class="sxs-lookup"><span data-stu-id="af492-219">Method</span></span> |

### <a name="example"></a><span data-ttu-id="af492-220">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-220">Example</span></span>

<span data-ttu-id="af492-221">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="af492-221">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="af492-222">Membros</span><span class="sxs-lookup"><span data-stu-id="af492-222">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="af492-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="af492-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="af492-224">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="af492-224">Gets the item's attachments as an array.</span></span> <span data-ttu-id="af492-225">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="af492-225">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-226">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="af492-226">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="af492-227">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="af492-227">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="af492-228">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-228">Type:</span></span>

*   <span data-ttu-id="af492-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="af492-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-230">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-230">Requirements</span></span>

|<span data-ttu-id="af492-231">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-231">Requirement</span></span>|<span data-ttu-id="af492-232">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-233">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-234">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-234">1.0</span></span>|
|[<span data-ttu-id="af492-235">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-235">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-236">ReadItem</span></span>|
|[<span data-ttu-id="af492-237">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-237">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-238">Read</span><span class="sxs-lookup"><span data-stu-id="af492-238">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-239">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-239">Example</span></span>

<span data-ttu-id="af492-240">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="af492-240">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="af492-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="af492-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="af492-242">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="af492-242">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="af492-243">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="af492-243">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-244">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-244">Type:</span></span>

*   [<span data-ttu-id="af492-245">Destinatários</span><span class="sxs-lookup"><span data-stu-id="af492-245">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="af492-246">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-246">Requirements</span></span>

|<span data-ttu-id="af492-247">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-247">Requirement</span></span>|<span data-ttu-id="af492-248">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-249">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-250">1.1</span><span class="sxs-lookup"><span data-stu-id="af492-250">1.1</span></span>|
|[<span data-ttu-id="af492-251">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-252">ReadItem</span></span>|
|[<span data-ttu-id="af492-253">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-254">Escrever</span><span class="sxs-lookup"><span data-stu-id="af492-254">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-255">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-255">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="af492-256">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="af492-256">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="af492-257">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="af492-257">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-258">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-258">Type:</span></span>

*   [<span data-ttu-id="af492-259">Corpo</span><span class="sxs-lookup"><span data-stu-id="af492-259">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="af492-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-260">Requirements</span></span>

|<span data-ttu-id="af492-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-261">Requirement</span></span>|<span data-ttu-id="af492-262">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-264">1.1</span><span class="sxs-lookup"><span data-stu-id="af492-264">1.1</span></span>|
|[<span data-ttu-id="af492-265">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-265">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-266">ReadItem</span></span>|
|[<span data-ttu-id="af492-267">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-267">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-268">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-268">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="af492-269">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="af492-269">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="af492-270">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="af492-270">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="af492-271">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="af492-271">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="af492-272">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="af492-272">Read mode</span></span>

<span data-ttu-id="af492-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="af492-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="af492-275">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="af492-275">Compose mode</span></span>

<span data-ttu-id="af492-276">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="af492-276">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-277">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-277">Type:</span></span>

*   <span data-ttu-id="af492-278">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="af492-278">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-279">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-279">Requirements</span></span>

|<span data-ttu-id="af492-280">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-280">Requirement</span></span>|<span data-ttu-id="af492-281">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-282">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-283">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-283">1.0</span></span>|
|[<span data-ttu-id="af492-284">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-284">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-285">ReadItem</span></span>|
|[<span data-ttu-id="af492-286">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-286">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-287">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-287">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-288">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-288">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="af492-289">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="af492-289">(nullable) conversationId :String</span></span>

<span data-ttu-id="af492-290">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="af492-290">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="af492-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="af492-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="af492-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="af492-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-295">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-295">Type:</span></span>

*   <span data-ttu-id="af492-296">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="af492-296">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-297">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-297">Requirements</span></span>

|<span data-ttu-id="af492-298">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-298">Requirement</span></span>|<span data-ttu-id="af492-299">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-300">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-301">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-301">1.0</span></span>|
|[<span data-ttu-id="af492-302">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-302">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-303">ReadItem</span></span>|
|[<span data-ttu-id="af492-304">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-304">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-305">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-305">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="af492-306">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="af492-306">dateTimeCreated :Date</span></span>

<span data-ttu-id="af492-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="af492-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-309">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-309">Type:</span></span>

*   <span data-ttu-id="af492-310">Data</span><span class="sxs-lookup"><span data-stu-id="af492-310">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-311">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-311">Requirements</span></span>

|<span data-ttu-id="af492-312">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-312">Requirement</span></span>|<span data-ttu-id="af492-313">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-313">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-314">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-314">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-315">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-315">1.0</span></span>|
|[<span data-ttu-id="af492-316">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-316">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-317">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-317">ReadItem</span></span>|
|[<span data-ttu-id="af492-318">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-318">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-319">Read</span><span class="sxs-lookup"><span data-stu-id="af492-319">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-320">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-320">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="af492-321">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="af492-321">dateTimeModified :Date</span></span>

<span data-ttu-id="af492-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="af492-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-324">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="af492-324">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-325">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-325">Type:</span></span>

*   <span data-ttu-id="af492-326">Data</span><span class="sxs-lookup"><span data-stu-id="af492-326">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-327">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-327">Requirements</span></span>

|<span data-ttu-id="af492-328">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-328">Requirement</span></span>|<span data-ttu-id="af492-329">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-329">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-330">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-330">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-331">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-331">1.0</span></span>|
|[<span data-ttu-id="af492-332">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-332">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-333">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-333">ReadItem</span></span>|
|[<span data-ttu-id="af492-334">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-334">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-335">Read</span><span class="sxs-lookup"><span data-stu-id="af492-335">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-336">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-336">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="af492-337">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="af492-337">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="af492-338">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="af492-338">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="af492-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="af492-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="af492-341">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="af492-341">Read mode</span></span>

<span data-ttu-id="af492-342">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="af492-342">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="af492-343">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="af492-343">Compose mode</span></span>

<span data-ttu-id="af492-344">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="af492-344">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="af492-345">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="af492-345">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-346">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-346">Type:</span></span>

*   <span data-ttu-id="af492-347">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="af492-347">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-348">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-348">Requirements</span></span>

|<span data-ttu-id="af492-349">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-349">Requirement</span></span>|<span data-ttu-id="af492-350">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-351">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-352">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-352">1.0</span></span>|
|[<span data-ttu-id="af492-353">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-353">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-354">ReadItem</span></span>|
|[<span data-ttu-id="af492-355">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-355">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-356">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-356">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-357">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-357">Example</span></span>

<span data-ttu-id="af492-358">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="af492-358">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="af492-359">enhancedLocation:[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="af492-359">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="af492-360">Obtém ou define os locais de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="af492-360">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="af492-361">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="af492-361">Read mode</span></span>

<span data-ttu-id="af492-362">O `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que permite que você obtenha o conjunto de locais (cada um representado por um objeto [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associado ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="af492-362">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="af492-363">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="af492-363">Compose mode</span></span>

<span data-ttu-id="af492-364">O `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que fornece os métodos para obter, remover ou adicionar locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="af492-364">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-365">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-365">Type:</span></span>

*   [<span data-ttu-id="af492-366">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="af492-366">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="af492-367">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-367">Requirements</span></span>

|<span data-ttu-id="af492-368">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-368">Requirement</span></span>|<span data-ttu-id="af492-369">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-369">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-370">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-371">Visualização</span><span class="sxs-lookup"><span data-stu-id="af492-371">Preview</span></span>|
|[<span data-ttu-id="af492-372">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-372">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-373">ReadItem</span></span>|
|[<span data-ttu-id="af492-374">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-374">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-375">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-375">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-376">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-376">Example</span></span>

<span data-ttu-id="af492-377">O exemplo a seguir obtém os locais atuais associados ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="af492-377">The following example gets the current locations associated with the appointment.</span></span>

```javascript
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type == Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}

// Sample output:
// Display name: Conf Room 14
// Type: room
// Email address: cr14@contoso.com
// Display name: Paris
// Type: custom
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="af492-378">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="af492-378">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="af492-379">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="af492-379">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="af492-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="af492-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-382">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="af492-382">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="af492-383">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="af492-383">Read mode</span></span>

<span data-ttu-id="af492-384">A propriedade `from` retorna um objeto `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="af492-384">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="af492-385">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="af492-385">Compose mode</span></span>

<span data-ttu-id="af492-386">A propriedade `from` retorna um objeto `From` que fornece um método para obtenção do valor de from.</span><span class="sxs-lookup"><span data-stu-id="af492-386">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="af492-387">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-387">Type:</span></span>

*   <span data-ttu-id="af492-388">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="af492-388">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-389">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-389">Requirements</span></span>

|<span data-ttu-id="af492-390">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-390">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="af492-391">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-392">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-392">1.0</span></span>|<span data-ttu-id="af492-393">1.7</span><span class="sxs-lookup"><span data-stu-id="af492-393">1.7</span></span>|
|[<span data-ttu-id="af492-394">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-394">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-395">ReadItem</span></span>|<span data-ttu-id="af492-396">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="af492-396">ReadWriteItem</span></span>|
|[<span data-ttu-id="af492-397">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-397">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-398">Read</span><span class="sxs-lookup"><span data-stu-id="af492-398">Read</span></span>|<span data-ttu-id="af492-399">Escrever</span><span class="sxs-lookup"><span data-stu-id="af492-399">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="af492-400">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="af492-400">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="af492-401">Obtém ou define os cabeçalhos de internet de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="af492-401">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-402">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-402">Type:</span></span>

*   [<span data-ttu-id="af492-403">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="af492-403">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="af492-404">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-404">Requirements</span></span>

|<span data-ttu-id="af492-405">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-405">Requirement</span></span>|<span data-ttu-id="af492-406">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-407">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-408">Visualização</span><span class="sxs-lookup"><span data-stu-id="af492-408">Preview</span></span>|
|[<span data-ttu-id="af492-409">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-410">ReadItem</span></span>|
|[<span data-ttu-id="af492-411">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-412">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-412">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="af492-413">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="af492-413">internetMessageId :String</span></span>

<span data-ttu-id="af492-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="af492-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-416">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-416">Type:</span></span>

*   <span data-ttu-id="af492-417">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="af492-417">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-418">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-418">Requirements</span></span>

|<span data-ttu-id="af492-419">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-419">Requirement</span></span>|<span data-ttu-id="af492-420">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-421">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-422">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-422">1.0</span></span>|
|[<span data-ttu-id="af492-423">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-424">ReadItem</span></span>|
|[<span data-ttu-id="af492-425">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-426">Leitura</span><span class="sxs-lookup"><span data-stu-id="af492-426">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-427">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-427">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="af492-428">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="af492-428">itemClass :String</span></span>

<span data-ttu-id="af492-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="af492-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="af492-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="af492-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="af492-433">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-433">Type</span></span>|<span data-ttu-id="af492-434">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-434">Description</span></span>|<span data-ttu-id="af492-435">classe de item</span><span class="sxs-lookup"><span data-stu-id="af492-435">item class</span></span>|
|---|---|---|
|<span data-ttu-id="af492-436">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="af492-436">Appointment items</span></span>|<span data-ttu-id="af492-437">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="af492-437">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="af492-438">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="af492-438">Message items</span></span>|<span data-ttu-id="af492-439">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="af492-439">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="af492-440">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="af492-440">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-441">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-441">Type:</span></span>

*   <span data-ttu-id="af492-442">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="af492-442">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-443">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-443">Requirements</span></span>

|<span data-ttu-id="af492-444">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-444">Requirement</span></span>|<span data-ttu-id="af492-445">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-446">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-447">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-447">1.0</span></span>|
|[<span data-ttu-id="af492-448">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-449">ReadItem</span></span>|
|[<span data-ttu-id="af492-450">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-451">Leitura</span><span class="sxs-lookup"><span data-stu-id="af492-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-452">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-452">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="af492-453">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="af492-453">(nullable) itemId :String</span></span>

<span data-ttu-id="af492-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="af492-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-456">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="af492-456">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="af492-457">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="af492-457">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="af492-458">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="af492-458">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="af492-459">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="af492-459">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="af492-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-462">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-462">Type:</span></span>

*   <span data-ttu-id="af492-463">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="af492-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-464">Requirements</span></span>

|<span data-ttu-id="af492-465">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-465">Requirement</span></span>|<span data-ttu-id="af492-466">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-467">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-468">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-468">1.0</span></span>|
|[<span data-ttu-id="af492-469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-470">ReadItem</span></span>|
|[<span data-ttu-id="af492-471">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-472">Leitura</span><span class="sxs-lookup"><span data-stu-id="af492-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-473">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-473">Example</span></span>

<span data-ttu-id="af492-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="af492-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="af492-476">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="af492-476">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="af492-477">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="af492-477">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="af492-478">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="af492-478">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-479">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-479">Type:</span></span>

*   [<span data-ttu-id="af492-480">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="af492-480">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="af492-481">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-481">Requirements</span></span>

|<span data-ttu-id="af492-482">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-482">Requirement</span></span>|<span data-ttu-id="af492-483">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-484">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-485">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-485">1.0</span></span>|
|[<span data-ttu-id="af492-486">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-486">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-487">ReadItem</span></span>|
|[<span data-ttu-id="af492-488">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-488">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-489">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-489">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-490">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-490">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="af492-491">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="af492-491">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="af492-492">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="af492-492">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="af492-493">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="af492-493">Read mode</span></span>

<span data-ttu-id="af492-494">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="af492-494">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="af492-495">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="af492-495">Compose mode</span></span>

<span data-ttu-id="af492-496">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="af492-496">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-497">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-497">Type:</span></span>

*   <span data-ttu-id="af492-498">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="af492-498">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-499">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-499">Requirements</span></span>

|<span data-ttu-id="af492-500">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-500">Requirement</span></span>|<span data-ttu-id="af492-501">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-502">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-503">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-503">1.0</span></span>|
|[<span data-ttu-id="af492-504">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-505">ReadItem</span></span>|
|[<span data-ttu-id="af492-506">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-507">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-507">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-508">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-508">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="af492-509">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="af492-509">normalizedSubject :String</span></span>

<span data-ttu-id="af492-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="af492-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="af492-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="af492-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-514">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-514">Type:</span></span>

*   <span data-ttu-id="af492-515">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="af492-515">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-516">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-516">Requirements</span></span>

|<span data-ttu-id="af492-517">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-517">Requirement</span></span>|<span data-ttu-id="af492-518">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-519">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-520">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-520">1.0</span></span>|
|[<span data-ttu-id="af492-521">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-521">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-522">ReadItem</span></span>|
|[<span data-ttu-id="af492-523">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-523">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-524">Leitura</span><span class="sxs-lookup"><span data-stu-id="af492-524">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-525">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-525">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="af492-526">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="af492-526">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="af492-527">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="af492-527">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-528">Type:</span><span class="sxs-lookup"><span data-stu-id="af492-528">Type:</span></span>

*   [<span data-ttu-id="af492-529">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="af492-529">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="af492-530">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-530">Requirements</span></span>

|<span data-ttu-id="af492-531">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-531">Requirement</span></span>|<span data-ttu-id="af492-532">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-532">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-533">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-533">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-534">1.3</span><span class="sxs-lookup"><span data-stu-id="af492-534">1.3</span></span>|
|[<span data-ttu-id="af492-535">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-535">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-536">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-536">ReadItem</span></span>|
|[<span data-ttu-id="af492-537">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-537">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-538">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-538">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="af492-539">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="af492-539">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="af492-540">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="af492-540">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="af492-541">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="af492-541">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="af492-542">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="af492-542">Read mode</span></span>

<span data-ttu-id="af492-543">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="af492-543">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="af492-544">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="af492-544">Compose mode</span></span>

<span data-ttu-id="af492-545">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="af492-545">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-546">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-546">Type:</span></span>

*   <span data-ttu-id="af492-547">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="af492-547">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-548">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-548">Requirements</span></span>

|<span data-ttu-id="af492-549">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-549">Requirement</span></span>|<span data-ttu-id="af492-550">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-551">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-552">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-552">1.0</span></span>|
|[<span data-ttu-id="af492-553">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-554">ReadItem</span></span>|
|[<span data-ttu-id="af492-555">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-556">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-556">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-557">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-557">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="af492-558">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="af492-558">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="af492-559">Obtém o endereço de email do organizador para uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="af492-559">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="af492-560">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="af492-560">Read mode</span></span>

<span data-ttu-id="af492-561">A propriedade `organizer` retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="af492-561">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="af492-562">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="af492-562">Compose mode</span></span>

<span data-ttu-id="af492-563">A propriedade `organizer` retorna um objeto [Organizer](/javascript/api/outlook/office.organizer) que fornece um método para obtenção do valor de organizer.</span><span class="sxs-lookup"><span data-stu-id="af492-563">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-564">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-564">Type:</span></span>

*   <span data-ttu-id="af492-565">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="af492-565">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-566">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-566">Requirements</span></span>

|<span data-ttu-id="af492-567">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-567">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="af492-568">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-568">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-569">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-569">1.0</span></span>|<span data-ttu-id="af492-570">1.7</span><span class="sxs-lookup"><span data-stu-id="af492-570">1.7</span></span>|
|[<span data-ttu-id="af492-571">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-572">ReadItem</span></span>|<span data-ttu-id="af492-573">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="af492-573">ReadWriteItem</span></span>|
|[<span data-ttu-id="af492-574">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-575">Read</span><span class="sxs-lookup"><span data-stu-id="af492-575">Read</span></span>|<span data-ttu-id="af492-576">Escrever</span><span class="sxs-lookup"><span data-stu-id="af492-576">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-577">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-577">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="af492-578">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="af492-578">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="af492-579">Obtém ou configura o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="af492-579">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="af492-580">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="af492-580">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="af492-581">Modos de leitura e redação para itens do compromisso.</span><span class="sxs-lookup"><span data-stu-id="af492-581">Read and compose modes for appointment items.</span></span> <span data-ttu-id="af492-582">Modo de leitura para os itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="af492-582">Read mode for meeting request items.</span></span>

<span data-ttu-id="af492-583">A propriedade `recurrence` retorna um objeto [recurrence](/javascript/api/outlook/office.recurrence) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="af492-583">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="af492-584">`null` retorna para compromissos individuais e solicitações de reunião de compromissos individuais.</span><span class="sxs-lookup"><span data-stu-id="af492-584">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="af492-585">`undefined` retorna para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="af492-585">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="af492-586">Observação: solicitações de reunião têm um valor `itemClass` de IPM.Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="af492-586">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="af492-587">Observação: se o objeto de recorrência for `null`, isso indicará que o objeto é um compromisso individual ou uma solicitação de reunião de um compromisso individual e NÃO parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="af492-587">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-588">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-588">Type:</span></span>

* [<span data-ttu-id="af492-589">Recurrence</span><span class="sxs-lookup"><span data-stu-id="af492-589">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="af492-590">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-590">Requirement</span></span>|<span data-ttu-id="af492-591">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-592">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-593">1.7</span><span class="sxs-lookup"><span data-stu-id="af492-593">1.7</span></span>|
|[<span data-ttu-id="af492-594">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-594">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-595">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-595">ReadItem</span></span>|
|[<span data-ttu-id="af492-596">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-596">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-597">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-597">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="af492-598">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="af492-598">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="af492-599">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="af492-599">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="af492-600">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="af492-600">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="af492-601">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="af492-601">Read mode</span></span>

<span data-ttu-id="af492-602">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="af492-602">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="af492-603">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="af492-603">Compose mode</span></span>

<span data-ttu-id="af492-604">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="af492-604">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-605">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-605">Type:</span></span>

*   <span data-ttu-id="af492-606">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="af492-606">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-607">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-607">Requirements</span></span>

|<span data-ttu-id="af492-608">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-608">Requirement</span></span>|<span data-ttu-id="af492-609">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-609">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-610">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-610">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-611">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-611">1.0</span></span>|
|[<span data-ttu-id="af492-612">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-612">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-613">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-613">ReadItem</span></span>|
|[<span data-ttu-id="af492-614">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-614">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-615">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-615">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-616">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-616">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="af492-617">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="af492-617">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="af492-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="af492-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="af492-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="af492-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-622">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="af492-622">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-623">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-623">Type:</span></span>

*   [<span data-ttu-id="af492-624">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="af492-624">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="af492-625">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-625">Requirements</span></span>

|<span data-ttu-id="af492-626">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-626">Requirement</span></span>|<span data-ttu-id="af492-627">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-628">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-628">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-629">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-629">1.0</span></span>|
|[<span data-ttu-id="af492-630">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-631">ReadItem</span></span>|
|[<span data-ttu-id="af492-632">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-633">Read</span><span class="sxs-lookup"><span data-stu-id="af492-633">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-634">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-634">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="af492-635">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="af492-635">(nullable) seriesId :String</span></span>

<span data-ttu-id="af492-636">Obtém a id da série a qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="af492-636">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="af492-637">No OWA e no Outlook, o `seriesId` retorna a ID dos Serviços Web do Exchange (EWS) do item pai (série) a qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="af492-637">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="af492-638">No entanto, no iOS e no Android, o `seriesId` retorna a ID REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="af492-638">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-639">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="af492-639">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="af492-640">A propriedade `seriesId` não é idêntica à ID do Outlook usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="af492-640">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="af492-641">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="af492-641">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="af492-642">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="af492-642">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="af492-643">A propriedade `seriesId` retorna `null` para itens que não têm itens pai como compromissos individuais, itens de série ou solicitações de reunião e retorna `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="af492-643">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-644">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-644">Type:</span></span>

* <span data-ttu-id="af492-645">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="af492-645">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-646">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-646">Requirements</span></span>

|<span data-ttu-id="af492-647">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-647">Requirement</span></span>|<span data-ttu-id="af492-648">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-648">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-649">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-649">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-650">1.7</span><span class="sxs-lookup"><span data-stu-id="af492-650">1.7</span></span>|
|[<span data-ttu-id="af492-651">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-651">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-652">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-652">ReadItem</span></span>|
|[<span data-ttu-id="af492-653">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-653">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-654">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-654">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-655">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-655">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="af492-656">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="af492-656">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="af492-657">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="af492-657">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="af492-p130">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="af492-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="af492-660">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="af492-660">Read mode</span></span>

<span data-ttu-id="af492-661">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="af492-661">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="af492-662">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="af492-662">Compose mode</span></span>

<span data-ttu-id="af492-663">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="af492-663">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="af492-664">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="af492-664">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-665">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-665">Type:</span></span>

*   <span data-ttu-id="af492-666">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="af492-666">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-667">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-667">Requirements</span></span>

|<span data-ttu-id="af492-668">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-668">Requirement</span></span>|<span data-ttu-id="af492-669">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-670">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-671">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-671">1.0</span></span>|
|[<span data-ttu-id="af492-672">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-672">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-673">ReadItem</span></span>|
|[<span data-ttu-id="af492-674">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-674">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-675">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-675">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-676">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-676">Example</span></span>

<span data-ttu-id="af492-677">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="af492-677">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="af492-678">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="af492-678">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="af492-679">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="af492-679">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="af492-680">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="af492-680">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="af492-681">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="af492-681">Read mode</span></span>

<span data-ttu-id="af492-p131">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="af492-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="af492-684">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="af492-684">Compose mode</span></span>

<span data-ttu-id="af492-685">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="af492-685">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="af492-686">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-686">Type:</span></span>

*   <span data-ttu-id="af492-687">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="af492-687">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-688">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-688">Requirements</span></span>

|<span data-ttu-id="af492-689">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-689">Requirement</span></span>|<span data-ttu-id="af492-690">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-691">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-692">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-692">1.0</span></span>|
|[<span data-ttu-id="af492-693">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-693">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-694">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-694">ReadItem</span></span>|
|[<span data-ttu-id="af492-695">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-695">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-696">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-696">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="af492-697">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="af492-697">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="af492-698">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="af492-698">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="af492-699">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="af492-699">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="af492-700">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="af492-700">Read mode</span></span>

<span data-ttu-id="af492-p133">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="af492-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="af492-703">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="af492-703">Compose mode</span></span>

<span data-ttu-id="af492-704">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="af492-704">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="af492-705">Tipo:</span><span class="sxs-lookup"><span data-stu-id="af492-705">Type:</span></span>

*   <span data-ttu-id="af492-706">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="af492-706">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-707">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-707">Requirements</span></span>

|<span data-ttu-id="af492-708">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-708">Requirement</span></span>|<span data-ttu-id="af492-709">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-709">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-710">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-710">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-711">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-711">1.0</span></span>|
|[<span data-ttu-id="af492-712">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-712">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-713">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-713">ReadItem</span></span>|
|[<span data-ttu-id="af492-714">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-714">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-715">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-715">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-716">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-716">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="af492-717">Métodos</span><span class="sxs-lookup"><span data-stu-id="af492-717">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="af492-718">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="af492-718">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="af492-719">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="af492-719">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="af492-720">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="af492-720">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="af492-721">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="af492-721">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-722">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-722">Parameters:</span></span>
|<span data-ttu-id="af492-723">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-723">Name</span></span>|<span data-ttu-id="af492-724">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-724">Type</span></span>|<span data-ttu-id="af492-725">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-725">Attributes</span></span>|<span data-ttu-id="af492-726">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-726">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="af492-727">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="af492-727">String</span></span>||<span data-ttu-id="af492-p134">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="af492-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="af492-730">String</span><span class="sxs-lookup"><span data-stu-id="af492-730">String</span></span>||<span data-ttu-id="af492-p135">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="af492-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="af492-733">Object</span><span class="sxs-lookup"><span data-stu-id="af492-733">Object</span></span>|<span data-ttu-id="af492-734">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-734">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-735">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-735">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="af492-736">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-736">Object</span></span>|<span data-ttu-id="af492-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-737">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-738">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-738">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="af492-739">Booliano</span><span class="sxs-lookup"><span data-stu-id="af492-739">Boolean</span></span>|<span data-ttu-id="af492-740">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-740">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-741">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="af492-741">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="af492-742">function</span><span class="sxs-lookup"><span data-stu-id="af492-742">function</span></span>|<span data-ttu-id="af492-743">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-743">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-744">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="af492-745">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="af492-745">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="af492-746">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="af492-746">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="af492-747">Erros</span><span class="sxs-lookup"><span data-stu-id="af492-747">Errors</span></span>

|<span data-ttu-id="af492-748">Código de erro</span><span class="sxs-lookup"><span data-stu-id="af492-748">Error code</span></span>|<span data-ttu-id="af492-749">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-749">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="af492-750">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="af492-750">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="af492-751">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="af492-751">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="af492-752">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="af492-752">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-753">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-753">Requirements</span></span>

|<span data-ttu-id="af492-754">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-754">Requirement</span></span>|<span data-ttu-id="af492-755">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-755">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-756">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-756">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-757">1.1</span><span class="sxs-lookup"><span data-stu-id="af492-757">1.1</span></span>|
|[<span data-ttu-id="af492-758">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-758">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-759">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="af492-759">ReadWriteItem</span></span>|
|[<span data-ttu-id="af492-760">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-760">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-761">Escrever</span><span class="sxs-lookup"><span data-stu-id="af492-761">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="af492-762">Exemplos</span><span class="sxs-lookup"><span data-stu-id="af492-762">Examples</span></span>

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

<span data-ttu-id="af492-763">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="af492-763">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="af492-764">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="af492-764">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="af492-765">Adiciona um arquivo a partir da codificação base64 a uma mensagem ou compromisso como anexo.</span><span class="sxs-lookup"><span data-stu-id="af492-765">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="af492-766">O método `addFileAttachmentFromBase64Async` carrega o arquivo a partir da codificação base64 e o anexa ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="af492-766">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="af492-767">Esse método retorna o identificador de anexo no objeto AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="af492-767">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="af492-768">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="af492-768">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-769">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-769">Parameters:</span></span>
|<span data-ttu-id="af492-770">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-770">Name</span></span>|<span data-ttu-id="af492-771">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-771">Type</span></span>|<span data-ttu-id="af492-772">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-772">Attributes</span></span>|<span data-ttu-id="af492-773">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-773">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="af492-774">String</span><span class="sxs-lookup"><span data-stu-id="af492-774">String</span></span>||<span data-ttu-id="af492-775">O conteúdo codificado em Base 64 de uma imagem ou arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="af492-775">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="af492-776">String</span><span class="sxs-lookup"><span data-stu-id="af492-776">String</span></span>||<span data-ttu-id="af492-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="af492-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="af492-779">Object</span><span class="sxs-lookup"><span data-stu-id="af492-779">Object</span></span>|<span data-ttu-id="af492-780">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-780">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-781">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-781">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="af492-782">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-782">Object</span></span>|<span data-ttu-id="af492-783">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-783">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-784">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-784">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="af492-785">Booliano</span><span class="sxs-lookup"><span data-stu-id="af492-785">Boolean</span></span>|<span data-ttu-id="af492-786">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-786">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-787">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="af492-787">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="af492-788">function</span><span class="sxs-lookup"><span data-stu-id="af492-788">function</span></span>|<span data-ttu-id="af492-789">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-789">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-790">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="af492-791">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="af492-791">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="af492-792">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="af492-792">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="af492-793">Erros</span><span class="sxs-lookup"><span data-stu-id="af492-793">Errors</span></span>

|<span data-ttu-id="af492-794">Código de erro</span><span class="sxs-lookup"><span data-stu-id="af492-794">Error code</span></span>|<span data-ttu-id="af492-795">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-795">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="af492-796">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="af492-796">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="af492-797">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="af492-797">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="af492-798">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="af492-798">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-799">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-799">Requirements</span></span>

|<span data-ttu-id="af492-800">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-800">Requirement</span></span>|<span data-ttu-id="af492-801">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-801">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-802">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-802">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-803">Visualização</span><span class="sxs-lookup"><span data-stu-id="af492-803">Preview</span></span>|
|[<span data-ttu-id="af492-804">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-804">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-805">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="af492-805">ReadWriteItem</span></span>|
|[<span data-ttu-id="af492-806">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-806">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-807">Redação</span><span class="sxs-lookup"><span data-stu-id="af492-807">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="af492-808">Exemplos</span><span class="sxs-lookup"><span data-stu-id="af492-808">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="af492-809">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="af492-809">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="af492-810">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="af492-810">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="af492-811">Atualmente, os tipos de evento compatíveis são `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="af492-811">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-812">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-812">Parameters:</span></span>

| <span data-ttu-id="af492-813">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-813">Name</span></span> | <span data-ttu-id="af492-814">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-814">Type</span></span> | <span data-ttu-id="af492-815">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-815">Attributes</span></span> | <span data-ttu-id="af492-816">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-816">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="af492-817">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="af492-817">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="af492-818">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="af492-818">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="af492-819">Função</span><span class="sxs-lookup"><span data-stu-id="af492-819">Function</span></span> || <span data-ttu-id="af492-p138">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="af492-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="af492-823">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-823">Object</span></span> | <span data-ttu-id="af492-824">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-824">&lt;optional&gt;</span></span> | <span data-ttu-id="af492-825">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-825">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="af492-826">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-826">Object</span></span> | <span data-ttu-id="af492-827">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-827">&lt;optional&gt;</span></span> | <span data-ttu-id="af492-828">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-828">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="af492-829">function</span><span class="sxs-lookup"><span data-stu-id="af492-829">function</span></span>| <span data-ttu-id="af492-830">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-830">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-831">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-831">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-832">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-832">Requirements</span></span>

|<span data-ttu-id="af492-833">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-833">Requirement</span></span>| <span data-ttu-id="af492-834">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-835">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af492-836">1.7</span><span class="sxs-lookup"><span data-stu-id="af492-836">1.7</span></span> |
|[<span data-ttu-id="af492-837">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af492-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-838">ReadItem</span></span> |
|[<span data-ttu-id="af492-839">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="af492-840">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-840">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="af492-841">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="af492-841">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="af492-842">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="af492-842">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="af492-p139">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="af492-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="af492-846">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="af492-846">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="af492-847">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="af492-847">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-848">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-848">Parameters:</span></span>

|<span data-ttu-id="af492-849">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-849">Name</span></span>|<span data-ttu-id="af492-850">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-850">Type</span></span>|<span data-ttu-id="af492-851">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-851">Attributes</span></span>|<span data-ttu-id="af492-852">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-852">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="af492-853">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="af492-853">String</span></span>||<span data-ttu-id="af492-p140">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="af492-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="af492-856">String</span><span class="sxs-lookup"><span data-stu-id="af492-856">String</span></span>||<span data-ttu-id="af492-p141">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="af492-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="af492-859">Object</span><span class="sxs-lookup"><span data-stu-id="af492-859">Object</span></span>|<span data-ttu-id="af492-860">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-860">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-861">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="af492-862">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-862">Object</span></span>|<span data-ttu-id="af492-863">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-863">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-864">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="af492-865">function</span><span class="sxs-lookup"><span data-stu-id="af492-865">function</span></span>|<span data-ttu-id="af492-866">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-866">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-867">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="af492-868">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="af492-868">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="af492-869">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="af492-869">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="af492-870">Erros</span><span class="sxs-lookup"><span data-stu-id="af492-870">Errors</span></span>

|<span data-ttu-id="af492-871">Código de erro</span><span class="sxs-lookup"><span data-stu-id="af492-871">Error code</span></span>|<span data-ttu-id="af492-872">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-872">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="af492-873">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="af492-873">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-874">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-874">Requirements</span></span>

|<span data-ttu-id="af492-875">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-875">Requirement</span></span>|<span data-ttu-id="af492-876">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-876">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-877">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-877">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-878">1.1</span><span class="sxs-lookup"><span data-stu-id="af492-878">1.1</span></span>|
|[<span data-ttu-id="af492-879">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-879">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-880">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="af492-880">ReadWriteItem</span></span>|
|[<span data-ttu-id="af492-881">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-881">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-882">Escrever</span><span class="sxs-lookup"><span data-stu-id="af492-882">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-883">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-883">Example</span></span>

<span data-ttu-id="af492-884">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="af492-884">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="af492-885">close()</span><span class="sxs-lookup"><span data-stu-id="af492-885">close()</span></span>

<span data-ttu-id="af492-886">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="af492-886">Closes the current item that is being composed.</span></span>

<span data-ttu-id="af492-p142">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="af492-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-889">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="af492-889">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="af492-890">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="af492-890">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-891">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-891">Requirements</span></span>

|<span data-ttu-id="af492-892">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-892">Requirement</span></span>|<span data-ttu-id="af492-893">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-893">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-894">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-894">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-895">1.3</span><span class="sxs-lookup"><span data-stu-id="af492-895">1.3</span></span>|
|[<span data-ttu-id="af492-896">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-896">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-897">Restricted</span><span class="sxs-lookup"><span data-stu-id="af492-897">Restricted</span></span>|
|[<span data-ttu-id="af492-898">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-898">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-899">Escrever</span><span class="sxs-lookup"><span data-stu-id="af492-899">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="af492-900">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="af492-900">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="af492-901">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="af492-901">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-902">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="af492-902">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="af492-903">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="af492-903">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="af492-904">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="af492-904">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="af492-p143">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="af492-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-908">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-908">Parameters:</span></span>

|<span data-ttu-id="af492-909">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-909">Name</span></span>|<span data-ttu-id="af492-910">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-910">Type</span></span>|<span data-ttu-id="af492-911">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-911">Attributes</span></span>|<span data-ttu-id="af492-912">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="af492-913">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="af492-913">String &#124; Object</span></span>||<span data-ttu-id="af492-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="af492-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="af492-916">**OU**</span><span class="sxs-lookup"><span data-stu-id="af492-916">**OR**</span></span><br/><span data-ttu-id="af492-p145">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="af492-919">String</span><span class="sxs-lookup"><span data-stu-id="af492-919">String</span></span>|<span data-ttu-id="af492-920">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-920">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="af492-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="af492-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="af492-924">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-924">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-925">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="af492-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="af492-926">String</span><span class="sxs-lookup"><span data-stu-id="af492-926">String</span></span>||<span data-ttu-id="af492-p147">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="af492-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="af492-929">String</span><span class="sxs-lookup"><span data-stu-id="af492-929">String</span></span>||<span data-ttu-id="af492-930">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="af492-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="af492-931">String</span><span class="sxs-lookup"><span data-stu-id="af492-931">String</span></span>||<span data-ttu-id="af492-p148">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="af492-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="af492-934">Booliano</span><span class="sxs-lookup"><span data-stu-id="af492-934">Boolean</span></span>||<span data-ttu-id="af492-p149">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="af492-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="af492-937">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="af492-937">String</span></span>||<span data-ttu-id="af492-p150">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="af492-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="af492-941">function</span><span class="sxs-lookup"><span data-stu-id="af492-941">function</span></span>|<span data-ttu-id="af492-942">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-942">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-943">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-944">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-944">Requirements</span></span>

|<span data-ttu-id="af492-945">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-945">Requirement</span></span>|<span data-ttu-id="af492-946">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-947">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-948">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-948">1.0</span></span>|
|[<span data-ttu-id="af492-949">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-949">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-950">ReadItem</span></span>|
|[<span data-ttu-id="af492-951">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-951">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-952">Read</span><span class="sxs-lookup"><span data-stu-id="af492-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="af492-953">Exemplos</span><span class="sxs-lookup"><span data-stu-id="af492-953">Examples</span></span>

<span data-ttu-id="af492-954">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="af492-954">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="af492-955">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="af492-955">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="af492-956">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="af492-956">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="af492-957">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="af492-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="af492-958">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="af492-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="af492-959">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="af492-960">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="af492-960">displayReplyForm(formData)</span></span>

<span data-ttu-id="af492-961">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="af492-961">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-962">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="af492-962">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="af492-963">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="af492-963">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="af492-964">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="af492-964">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="af492-p151">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="af492-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-968">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-968">Parameters:</span></span>

|<span data-ttu-id="af492-969">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-969">Name</span></span>|<span data-ttu-id="af492-970">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-970">Type</span></span>|<span data-ttu-id="af492-971">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-971">Attributes</span></span>|<span data-ttu-id="af492-972">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-972">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="af492-973">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="af492-973">String &#124; Object</span></span>||<span data-ttu-id="af492-p152">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="af492-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="af492-976">**OU**</span><span class="sxs-lookup"><span data-stu-id="af492-976">**OR**</span></span><br/><span data-ttu-id="af492-p153">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="af492-979">String</span><span class="sxs-lookup"><span data-stu-id="af492-979">String</span></span>|<span data-ttu-id="af492-980">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-980">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="af492-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="af492-983">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-983">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="af492-984">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-984">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-985">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="af492-985">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="af492-986">String</span><span class="sxs-lookup"><span data-stu-id="af492-986">String</span></span>||<span data-ttu-id="af492-p155">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="af492-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="af492-989">String</span><span class="sxs-lookup"><span data-stu-id="af492-989">String</span></span>||<span data-ttu-id="af492-990">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="af492-990">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="af492-991">String</span><span class="sxs-lookup"><span data-stu-id="af492-991">String</span></span>||<span data-ttu-id="af492-p156">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="af492-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="af492-994">Booliano</span><span class="sxs-lookup"><span data-stu-id="af492-994">Boolean</span></span>||<span data-ttu-id="af492-p157">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="af492-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="af492-997">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="af492-997">String</span></span>||<span data-ttu-id="af492-p158">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="af492-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="af492-1001">function</span><span class="sxs-lookup"><span data-stu-id="af492-1001">function</span></span>|<span data-ttu-id="af492-1002">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1002">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1003">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-1003">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1004">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1004">Requirements</span></span>

|<span data-ttu-id="af492-1005">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1005">Requirement</span></span>|<span data-ttu-id="af492-1006">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1007">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1007">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-1008">1.0</span></span>|
|[<span data-ttu-id="af492-1009">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1010">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-1010">ReadItem</span></span>|
|[<span data-ttu-id="af492-1011">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1012">Read</span><span class="sxs-lookup"><span data-stu-id="af492-1012">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="af492-1013">Exemplos</span><span class="sxs-lookup"><span data-stu-id="af492-1013">Examples</span></span>

<span data-ttu-id="af492-1014">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="af492-1014">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="af492-1015">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="af492-1015">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="af492-1016">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="af492-1016">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="af492-1017">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="af492-1017">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="af492-1018">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="af492-1018">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="af492-1019">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-1019">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="af492-1020">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="af492-1020">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="af492-1021">Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um objeto `AttachmentContent`.</span><span class="sxs-lookup"><span data-stu-id="af492-1021">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="af492-1022">O método `getAttachmentContentAsync` remove o obtém anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="af492-1022">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="af492-1023">Como melhor prática, você deve usar o identificador para recuperar um anexo na mesma sessão da qual attachmentIds foram recuperadas com o chamada `getAttachmentsAsync` ou `item.attachments`.</span><span class="sxs-lookup"><span data-stu-id="af492-1023">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="af492-1024">No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="af492-1024">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="af492-1025">Uma sessão é finalizada quando o usuário fecha o aplicativo, ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="af492-1025">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-1026">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-1026">Parameters:</span></span>

|<span data-ttu-id="af492-1027">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-1027">Name</span></span>|<span data-ttu-id="af492-1028">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-1028">Type</span></span>|<span data-ttu-id="af492-1029">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-1029">Attributes</span></span>|<span data-ttu-id="af492-1030">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1030">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="af492-1031">String</span><span class="sxs-lookup"><span data-stu-id="af492-1031">String</span></span>||<span data-ttu-id="af492-1032">O identificador do anexo que você quer obter.</span><span class="sxs-lookup"><span data-stu-id="af492-1032">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="af492-1033">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1033">Object</span></span>|<span data-ttu-id="af492-1034">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1034">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1035">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-1035">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="af492-1036">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1036">Object</span></span>|<span data-ttu-id="af492-1037">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1038">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-1038">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="af492-1039">function</span><span class="sxs-lookup"><span data-stu-id="af492-1039">function</span></span>|<span data-ttu-id="af492-1040">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1041">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-1041">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1042">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1042">Requirements</span></span>

|<span data-ttu-id="af492-1043">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1043">Requirement</span></span>|<span data-ttu-id="af492-1044">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1044">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1045">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1045">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1046">Visualização</span><span class="sxs-lookup"><span data-stu-id="af492-1046">Preview</span></span>|
|[<span data-ttu-id="af492-1047">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1047">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1048">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-1048">ReadItem</span></span>|
|[<span data-ttu-id="af492-1049">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1049">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1050">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-1050">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="af492-1051">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af492-1051">Returns:</span></span>

<span data-ttu-id="af492-1052">Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="af492-1052">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="af492-1053">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1053">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="af492-1054">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="af492-1054">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="af492-1055">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="af492-1055">Gets the item's attachments as an array.</span></span> <span data-ttu-id="af492-1056">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="af492-1056">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-1057">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-1057">Parameters:</span></span>

|<span data-ttu-id="af492-1058">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-1058">Name</span></span>|<span data-ttu-id="af492-1059">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-1059">Type</span></span>|<span data-ttu-id="af492-1060">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-1060">Attributes</span></span>|<span data-ttu-id="af492-1061">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1061">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="af492-1062">Object</span><span class="sxs-lookup"><span data-stu-id="af492-1062">Object</span></span>|<span data-ttu-id="af492-1063">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1064">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-1064">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="af492-1065">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1065">Object</span></span>|<span data-ttu-id="af492-1066">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1067">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-1067">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="af492-1068">function</span><span class="sxs-lookup"><span data-stu-id="af492-1068">function</span></span>|<span data-ttu-id="af492-1069">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1070">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-1070">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1071">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1071">Requirements</span></span>

|<span data-ttu-id="af492-1072">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1072">Requirement</span></span>|<span data-ttu-id="af492-1073">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1073">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1074">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1074">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1075">Visualização</span><span class="sxs-lookup"><span data-stu-id="af492-1075">Preview</span></span>|
|[<span data-ttu-id="af492-1076">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1076">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1077">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-1077">ReadItem</span></span>|
|[<span data-ttu-id="af492-1078">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1078">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1079">Escrever</span><span class="sxs-lookup"><span data-stu-id="af492-1079">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="af492-1080">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af492-1080">Returns:</span></span>

<span data-ttu-id="af492-1081">Tipo: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="af492-1081">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="af492-1082">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1082">Example</span></span>

<span data-ttu-id="af492-1083">O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="af492-1083">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="af492-1084">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="af492-1084">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="af492-1085">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="af492-1085">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-1086">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="af492-1086">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-1087">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1087">Requirements</span></span>

|<span data-ttu-id="af492-1088">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1088">Requirement</span></span>|<span data-ttu-id="af492-1089">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1090">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1090">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1091">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-1091">1.0</span></span>|
|[<span data-ttu-id="af492-1092">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1093">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-1093">ReadItem</span></span>|
|[<span data-ttu-id="af492-1094">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1095">Read</span><span class="sxs-lookup"><span data-stu-id="af492-1095">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="af492-1096">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af492-1096">Returns:</span></span>

<span data-ttu-id="af492-1097">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="af492-1097">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="af492-1098">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1098">Example</span></span>

<span data-ttu-id="af492-1099">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="af492-1099">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="af492-1100">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="af492-1100">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="af492-1101">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="af492-1101">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-1102">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="af492-1102">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-1103">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-1103">Parameters:</span></span>

|<span data-ttu-id="af492-1104">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-1104">Name</span></span>|<span data-ttu-id="af492-1105">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-1105">Type</span></span>|<span data-ttu-id="af492-1106">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1106">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="af492-1107">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="af492-1107">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="af492-1108">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="af492-1108">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1109">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1109">Requirements</span></span>

|<span data-ttu-id="af492-1110">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1110">Requirement</span></span>|<span data-ttu-id="af492-1111">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1111">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1112">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1112">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1113">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-1113">1.0</span></span>|
|[<span data-ttu-id="af492-1114">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1114">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1115">Restricted</span><span class="sxs-lookup"><span data-stu-id="af492-1115">Restricted</span></span>|
|[<span data-ttu-id="af492-1116">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1116">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1117">Read</span><span class="sxs-lookup"><span data-stu-id="af492-1117">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="af492-1118">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af492-1118">Returns:</span></span>

<span data-ttu-id="af492-1119">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="af492-1119">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="af492-1120">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="af492-1120">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="af492-1121">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="af492-1121">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="af492-1122">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-1122">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="af492-1123">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="af492-1123">Value of `entityType`</span></span>|<span data-ttu-id="af492-1124">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="af492-1124">Type of objects in returned array</span></span>|<span data-ttu-id="af492-1125">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="af492-1125">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="af492-1126">String</span><span class="sxs-lookup"><span data-stu-id="af492-1126">String</span></span>|<span data-ttu-id="af492-1127">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="af492-1127">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="af492-1128">Contato</span><span class="sxs-lookup"><span data-stu-id="af492-1128">Contact</span></span>|<span data-ttu-id="af492-1129">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="af492-1129">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="af492-1130">String</span><span class="sxs-lookup"><span data-stu-id="af492-1130">String</span></span>|<span data-ttu-id="af492-1131">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="af492-1131">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="af492-1132">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="af492-1132">MeetingSuggestion</span></span>|<span data-ttu-id="af492-1133">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="af492-1133">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="af492-1134">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="af492-1134">PhoneNumber</span></span>|<span data-ttu-id="af492-1135">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="af492-1135">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="af492-1136">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="af492-1136">TaskSuggestion</span></span>|<span data-ttu-id="af492-1137">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="af492-1137">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="af492-1138">String</span><span class="sxs-lookup"><span data-stu-id="af492-1138">String</span></span>|<span data-ttu-id="af492-1139">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="af492-1139">**Restricted**</span></span>|

<span data-ttu-id="af492-1140">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="af492-1140">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="af492-1141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1141">Example</span></span>

<span data-ttu-id="af492-1142">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="af492-1142">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="af492-1143">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="af492-1143">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="af492-1144">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="af492-1144">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-1145">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="af492-1145">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="af492-1146">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="af492-1146">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-1147">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-1147">Parameters:</span></span>

|<span data-ttu-id="af492-1148">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-1148">Name</span></span>|<span data-ttu-id="af492-1149">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-1149">Type</span></span>|<span data-ttu-id="af492-1150">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1150">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="af492-1151">String</span><span class="sxs-lookup"><span data-stu-id="af492-1151">String</span></span>|<span data-ttu-id="af492-1152">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="af492-1152">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1153">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1153">Requirements</span></span>

|<span data-ttu-id="af492-1154">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1154">Requirement</span></span>|<span data-ttu-id="af492-1155">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1155">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1156">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1157">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-1157">1.0</span></span>|
|[<span data-ttu-id="af492-1158">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1158">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1159">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-1159">ReadItem</span></span>|
|[<span data-ttu-id="af492-1160">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1160">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1161">Read</span><span class="sxs-lookup"><span data-stu-id="af492-1161">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="af492-1162">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af492-1162">Returns:</span></span>

<span data-ttu-id="af492-p162">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="af492-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="af492-1165">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="af492-1165">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="af492-1166">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="af492-1166">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="af492-1167">Obtém dados de inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="af492-1167">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="af492-1168">Esse método só é compatível com o Outlook 2016 ou posterior para Windows (versões Clique para Executar posteriores à 16.0.8413.1000) e o Outlook na Web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="af492-1168">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-1169">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-1169">Parameters:</span></span>
|<span data-ttu-id="af492-1170">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-1170">Name</span></span>|<span data-ttu-id="af492-1171">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-1171">Type</span></span>|<span data-ttu-id="af492-1172">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-1172">Attributes</span></span>|<span data-ttu-id="af492-1173">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1173">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="af492-1174">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1174">Object</span></span>|<span data-ttu-id="af492-1175">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1175">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1176">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-1176">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="af492-1177">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1177">Object</span></span>|<span data-ttu-id="af492-1178">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1178">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1179">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-1179">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="af492-1180">function</span><span class="sxs-lookup"><span data-stu-id="af492-1180">function</span></span>|<span data-ttu-id="af492-1181">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1181">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1182">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-1182">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="af492-1183">Após o êxito, os dados de inicialização são fornecidos na propriedade `asyncResult.value` como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="af492-1183">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="af492-1184">Se não houver nenhum contexto de inicialização, o objeto `asyncResult` conterá um objeto `Error` com sua propriedade `code` definida como `9020` e sua propriedade `name` definida como `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="af492-1184">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1185">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1185">Requirements</span></span>

|<span data-ttu-id="af492-1186">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1186">Requirement</span></span>|<span data-ttu-id="af492-1187">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1188">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1189">Visualização</span><span class="sxs-lookup"><span data-stu-id="af492-1189">Preview</span></span>|
|[<span data-ttu-id="af492-1190">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-1191">ReadItem</span></span>|
|[<span data-ttu-id="af492-1192">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1193">Leitura</span><span class="sxs-lookup"><span data-stu-id="af492-1193">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-1194">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1194">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="af492-1195">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="af492-1195">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="af492-1196">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="af492-1196">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-1197">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="af492-1197">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="af492-p163">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="af492-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="af492-1201">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="af492-1201">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="af492-1202">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="af492-1202">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="af492-p164">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="af492-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-1206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1206">Requirements</span></span>

|<span data-ttu-id="af492-1207">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1207">Requirement</span></span>|<span data-ttu-id="af492-1208">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1208">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1209">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1210">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-1210">1.0</span></span>|
|[<span data-ttu-id="af492-1211">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-1212">ReadItem</span></span>|
|[<span data-ttu-id="af492-1213">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1214">Read</span><span class="sxs-lookup"><span data-stu-id="af492-1214">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="af492-1215">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af492-1215">Returns:</span></span>

<span data-ttu-id="af492-p165">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="af492-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="af492-1218">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="af492-1218">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="af492-1219">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1219">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="af492-1220">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1220">Example</span></span>

<span data-ttu-id="af492-1221">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="af492-1221">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="af492-1222">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="af492-1222">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="af492-1223">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="af492-1223">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-1224">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="af492-1224">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="af492-1225">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="af492-1225">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="af492-p166">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="af492-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-1228">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-1228">Parameters:</span></span>

|<span data-ttu-id="af492-1229">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-1229">Name</span></span>|<span data-ttu-id="af492-1230">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-1230">Type</span></span>|<span data-ttu-id="af492-1231">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1231">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="af492-1232">String</span><span class="sxs-lookup"><span data-stu-id="af492-1232">String</span></span>|<span data-ttu-id="af492-1233">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="af492-1233">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1234">Requirements</span></span>

|<span data-ttu-id="af492-1235">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1235">Requirement</span></span>|<span data-ttu-id="af492-1236">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1236">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1238">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-1238">1.0</span></span>|
|[<span data-ttu-id="af492-1239">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-1240">ReadItem</span></span>|
|[<span data-ttu-id="af492-1241">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1242">Read</span><span class="sxs-lookup"><span data-stu-id="af492-1242">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="af492-1243">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af492-1243">Returns:</span></span>

<span data-ttu-id="af492-1244">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="af492-1244">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="af492-1245">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="af492-1245">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="af492-1246">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="af492-1246">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="af492-1247">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1247">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="af492-1248">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="af492-1248">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="af492-1249">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="af492-1249">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="af492-p167">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="af492-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-1252">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-1252">Parameters:</span></span>

|<span data-ttu-id="af492-1253">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-1253">Name</span></span>|<span data-ttu-id="af492-1254">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-1254">Type</span></span>|<span data-ttu-id="af492-1255">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-1255">Attributes</span></span>|<span data-ttu-id="af492-1256">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1256">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="af492-1257">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="af492-1257">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="af492-p168">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="af492-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="af492-1261">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1261">Object</span></span>|<span data-ttu-id="af492-1262">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1262">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1263">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-1263">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="af492-1264">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1264">Object</span></span>|<span data-ttu-id="af492-1265">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1266">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-1266">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="af492-1267">function</span><span class="sxs-lookup"><span data-stu-id="af492-1267">function</span></span>||<span data-ttu-id="af492-1268">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-1268">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="af492-1269">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="af492-1269">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="af492-1270">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="af492-1270">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1271">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1271">Requirements</span></span>

|<span data-ttu-id="af492-1272">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1272">Requirement</span></span>|<span data-ttu-id="af492-1273">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1273">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1274">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1275">1.2</span><span class="sxs-lookup"><span data-stu-id="af492-1275">1.2</span></span>|
|[<span data-ttu-id="af492-1276">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1277">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="af492-1277">ReadWriteItem</span></span>|
|[<span data-ttu-id="af492-1278">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1279">Escrever</span><span class="sxs-lookup"><span data-stu-id="af492-1279">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="af492-1280">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af492-1280">Returns:</span></span>

<span data-ttu-id="af492-1281">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="af492-1281">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="af492-1282">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="af492-1282">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="af492-1283">String</span><span class="sxs-lookup"><span data-stu-id="af492-1283">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="af492-1284">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1284">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="af492-1285">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="af492-1285">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="af492-p170">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="af492-p170">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="af492-1288">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="af492-1288">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-1289">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1289">Requirements</span></span>

|<span data-ttu-id="af492-1290">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1290">Requirement</span></span>|<span data-ttu-id="af492-1291">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1291">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1292">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1292">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1293">1.6</span><span class="sxs-lookup"><span data-stu-id="af492-1293">1.6</span></span>|
|[<span data-ttu-id="af492-1294">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1294">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1295">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-1295">ReadItem</span></span>|
|[<span data-ttu-id="af492-1296">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1296">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1297">Read</span><span class="sxs-lookup"><span data-stu-id="af492-1297">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="af492-1298">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af492-1298">Returns:</span></span>

<span data-ttu-id="af492-1299">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="af492-1299">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="af492-1300">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1300">Example</span></span>

<span data-ttu-id="af492-1301">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="af492-1301">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="af492-1302">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="af492-1302">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="af492-p171">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="af492-p171">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="af492-1305">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="af492-1305">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="af492-p172">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="af492-p172">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="af492-1309">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="af492-1309">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="af492-1310">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="af492-1310">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="af492-p173">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="af492-p173">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="af492-1314">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1314">Requirements</span></span>

|<span data-ttu-id="af492-1315">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1315">Requirement</span></span>|<span data-ttu-id="af492-1316">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1316">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1317">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1317">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1318">1.6</span><span class="sxs-lookup"><span data-stu-id="af492-1318">1.6</span></span>|
|[<span data-ttu-id="af492-1319">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1319">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1320">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-1320">ReadItem</span></span>|
|[<span data-ttu-id="af492-1321">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1321">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1322">Read</span><span class="sxs-lookup"><span data-stu-id="af492-1322">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="af492-1323">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af492-1323">Returns:</span></span>

<span data-ttu-id="af492-p174">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="af492-p174">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="af492-1326">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1326">Example</span></span>

<span data-ttu-id="af492-1327">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="af492-1327">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="af492-1328">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="af492-1328">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="af492-1329">Obtém as propriedades do compromisso ou mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="af492-1329">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-1330">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-1330">Parameters:</span></span>

|<span data-ttu-id="af492-1331">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-1331">Name</span></span>|<span data-ttu-id="af492-1332">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-1332">Type</span></span>|<span data-ttu-id="af492-1333">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-1333">Attributes</span></span>|<span data-ttu-id="af492-1334">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1334">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="af492-1335">Object</span><span class="sxs-lookup"><span data-stu-id="af492-1335">Object</span></span>|<span data-ttu-id="af492-1336">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1336">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1337">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-1337">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="af492-1338">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1338">Object</span></span>|<span data-ttu-id="af492-1339">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1339">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1340">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-1340">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="af492-1341">function</span><span class="sxs-lookup"><span data-stu-id="af492-1341">function</span></span>||<span data-ttu-id="af492-1342">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="af492-1343">As propriedades compartilhadas são fornecidas como um objeto [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="af492-1343">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="af492-1344">Esse objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="af492-1344">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1345">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1345">Requirements</span></span>

|<span data-ttu-id="af492-1346">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1346">Requirement</span></span>|<span data-ttu-id="af492-1347">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1347">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1348">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1349">Visualização</span><span class="sxs-lookup"><span data-stu-id="af492-1349">Preview</span></span>|
|[<span data-ttu-id="af492-1350">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-1351">ReadItem</span></span>|
|[<span data-ttu-id="af492-1352">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1353">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-1353">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-1354">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1354">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="af492-1355">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="af492-1355">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="af492-1356">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="af492-1356">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="af492-p176">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="af492-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-1360">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-1360">Parameters:</span></span>

|<span data-ttu-id="af492-1361">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-1361">Name</span></span>|<span data-ttu-id="af492-1362">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-1362">Type</span></span>|<span data-ttu-id="af492-1363">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-1363">Attributes</span></span>|<span data-ttu-id="af492-1364">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1364">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="af492-1365">function</span><span class="sxs-lookup"><span data-stu-id="af492-1365">function</span></span>||<span data-ttu-id="af492-1366">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-1366">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="af492-1367">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="af492-1367">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="af492-1368">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="af492-1368">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="af492-1369">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1369">Object</span></span>|<span data-ttu-id="af492-1370">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1370">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1371">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-1371">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="af492-1372">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-1372">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1373">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1373">Requirements</span></span>

|<span data-ttu-id="af492-1374">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1374">Requirement</span></span>|<span data-ttu-id="af492-1375">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1375">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1376">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1377">1.0</span><span class="sxs-lookup"><span data-stu-id="af492-1377">1.0</span></span>|
|[<span data-ttu-id="af492-1378">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-1379">ReadItem</span></span>|
|[<span data-ttu-id="af492-1380">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1381">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-1381">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-1382">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1382">Example</span></span>

<span data-ttu-id="af492-p179">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="af492-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="af492-1386">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="af492-1386">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="af492-1387">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="af492-1387">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="af492-1388">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="af492-1388">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="af492-1389">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="af492-1389">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="af492-1390">No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="af492-1390">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="af492-1391">Uma sessão é finalizada quando o usuário fecha o aplicativo, ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="af492-1391">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-1392">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-1392">Parameters:</span></span>

|<span data-ttu-id="af492-1393">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-1393">Name</span></span>|<span data-ttu-id="af492-1394">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-1394">Type</span></span>|<span data-ttu-id="af492-1395">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-1395">Attributes</span></span>|<span data-ttu-id="af492-1396">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1396">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="af492-1397">String</span><span class="sxs-lookup"><span data-stu-id="af492-1397">String</span></span>||<span data-ttu-id="af492-1398">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="af492-1398">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="af492-1399">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1399">Object</span></span>|<span data-ttu-id="af492-1400">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1400">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1401">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-1401">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="af492-1402">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1402">Object</span></span>|<span data-ttu-id="af492-1403">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1403">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1404">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-1404">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="af492-1405">function</span><span class="sxs-lookup"><span data-stu-id="af492-1405">function</span></span>|<span data-ttu-id="af492-1406">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1406">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1407">Quando o método for concluído, a função transmitida ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-1407">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="af492-1408">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="af492-1408">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="af492-1409">Erros</span><span class="sxs-lookup"><span data-stu-id="af492-1409">Errors</span></span>

|<span data-ttu-id="af492-1410">Código de erro</span><span class="sxs-lookup"><span data-stu-id="af492-1410">Error code</span></span>|<span data-ttu-id="af492-1411">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1411">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="af492-1412">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="af492-1412">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1413">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1413">Requirements</span></span>

|<span data-ttu-id="af492-1414">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1414">Requirement</span></span>|<span data-ttu-id="af492-1415">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1415">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1416">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1417">1.1</span><span class="sxs-lookup"><span data-stu-id="af492-1417">1.1</span></span>|
|[<span data-ttu-id="af492-1418">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1418">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1419">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="af492-1419">ReadWriteItem</span></span>|
|[<span data-ttu-id="af492-1420">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1420">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1421">Escrever</span><span class="sxs-lookup"><span data-stu-id="af492-1421">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-1422">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1422">Example</span></span>

<span data-ttu-id="af492-1423">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="af492-1423">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="af492-1424">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="af492-1424">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="af492-1425">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="af492-1425">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="af492-1426">Atualmente, os tipos de evento compatíveis são `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="af492-1426">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-1427">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-1427">Parameters:</span></span>

| <span data-ttu-id="af492-1428">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-1428">Name</span></span> | <span data-ttu-id="af492-1429">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-1429">Type</span></span> | <span data-ttu-id="af492-1430">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-1430">Attributes</span></span> | <span data-ttu-id="af492-1431">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1431">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="af492-1432">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="af492-1432">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="af492-1433">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="af492-1433">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="af492-1434">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1434">Object</span></span> | <span data-ttu-id="af492-1435">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1435">&lt;optional&gt;</span></span> | <span data-ttu-id="af492-1436">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-1436">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="af492-1437">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1437">Object</span></span> | <span data-ttu-id="af492-1438">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1438">&lt;optional&gt;</span></span> | <span data-ttu-id="af492-1439">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-1439">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="af492-1440">function</span><span class="sxs-lookup"><span data-stu-id="af492-1440">function</span></span>| <span data-ttu-id="af492-1441">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1441">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1442">Quando o método for concluído, a função transmitida ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-1442">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1443">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1443">Requirements</span></span>

|<span data-ttu-id="af492-1444">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1444">Requirement</span></span>| <span data-ttu-id="af492-1445">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1445">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1446">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af492-1447">1.7</span><span class="sxs-lookup"><span data-stu-id="af492-1447">1.7</span></span> |
|[<span data-ttu-id="af492-1448">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af492-1449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af492-1449">ReadItem</span></span> |
|[<span data-ttu-id="af492-1450">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="af492-1451">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="af492-1451">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="af492-1452">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="af492-1452">saveAsync([options], callback)</span></span>

<span data-ttu-id="af492-1453">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="af492-1453">Asynchronously saves an item.</span></span>

<span data-ttu-id="af492-p181">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="af492-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-1457">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="af492-1457">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="af492-1458">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="af492-1458">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="af492-p183">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="af492-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="af492-1462">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="af492-1462">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="af492-1463">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="af492-1463">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="af492-1464">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="af492-1464">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="af492-1465">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="af492-1465">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-1466">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-1466">Parameters:</span></span>

|<span data-ttu-id="af492-1467">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-1467">Name</span></span>|<span data-ttu-id="af492-1468">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-1468">Type</span></span>|<span data-ttu-id="af492-1469">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-1469">Attributes</span></span>|<span data-ttu-id="af492-1470">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1470">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="af492-1471">Object</span><span class="sxs-lookup"><span data-stu-id="af492-1471">Object</span></span>|<span data-ttu-id="af492-1472">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1472">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1473">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-1473">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="af492-1474">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1474">Object</span></span>|<span data-ttu-id="af492-1475">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1475">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1476">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-1476">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="af492-1477">function</span><span class="sxs-lookup"><span data-stu-id="af492-1477">function</span></span>||<span data-ttu-id="af492-1478">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-1478">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="af492-1479">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="af492-1479">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1480">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1480">Requirements</span></span>

|<span data-ttu-id="af492-1481">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1481">Requirement</span></span>|<span data-ttu-id="af492-1482">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1482">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1483">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1484">1.3</span><span class="sxs-lookup"><span data-stu-id="af492-1484">1.3</span></span>|
|[<span data-ttu-id="af492-1485">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1485">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1486">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="af492-1486">ReadWriteItem</span></span>|
|[<span data-ttu-id="af492-1487">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1487">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1488">Escrever</span><span class="sxs-lookup"><span data-stu-id="af492-1488">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="af492-1489">Exemplos</span><span class="sxs-lookup"><span data-stu-id="af492-1489">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="af492-p185">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="af492-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="af492-1492">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="af492-1492">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="af492-1493">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="af492-1493">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="af492-p186">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="af492-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af492-1497">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="af492-1497">Parameters:</span></span>

|<span data-ttu-id="af492-1498">Nome</span><span class="sxs-lookup"><span data-stu-id="af492-1498">Name</span></span>|<span data-ttu-id="af492-1499">Tipo</span><span class="sxs-lookup"><span data-stu-id="af492-1499">Type</span></span>|<span data-ttu-id="af492-1500">Atributos</span><span class="sxs-lookup"><span data-stu-id="af492-1500">Attributes</span></span>|<span data-ttu-id="af492-1501">Descrição</span><span class="sxs-lookup"><span data-stu-id="af492-1501">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="af492-1502">String</span><span class="sxs-lookup"><span data-stu-id="af492-1502">String</span></span>||<span data-ttu-id="af492-p187">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="af492-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="af492-1506">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1506">Object</span></span>|<span data-ttu-id="af492-1507">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1507">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1508">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af492-1508">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="af492-1509">Objeto</span><span class="sxs-lookup"><span data-stu-id="af492-1509">Object</span></span>|<span data-ttu-id="af492-1510">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1510">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-1511">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af492-1511">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="af492-1512">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="af492-1512">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="af492-1513">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af492-1513">&lt;optional&gt;</span></span>|<span data-ttu-id="af492-p188">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="af492-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="af492-p189">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="af492-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="af492-1518">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="af492-1518">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="af492-1519">function</span><span class="sxs-lookup"><span data-stu-id="af492-1519">function</span></span>||<span data-ttu-id="af492-1520">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af492-1520">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af492-1521">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af492-1521">Requirements</span></span>

|<span data-ttu-id="af492-1522">Requisito</span><span class="sxs-lookup"><span data-stu-id="af492-1522">Requirement</span></span>|<span data-ttu-id="af492-1523">Valor</span><span class="sxs-lookup"><span data-stu-id="af492-1523">Value</span></span>|
|---|---|
|[<span data-ttu-id="af492-1524">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af492-1524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="af492-1525">1.2</span><span class="sxs-lookup"><span data-stu-id="af492-1525">1.2</span></span>|
|[<span data-ttu-id="af492-1526">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af492-1526">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="af492-1527">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="af492-1527">ReadWriteItem</span></span>|
|[<span data-ttu-id="af492-1528">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af492-1528">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="af492-1529">Escrever</span><span class="sxs-lookup"><span data-stu-id="af492-1529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="af492-1530">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af492-1530">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
