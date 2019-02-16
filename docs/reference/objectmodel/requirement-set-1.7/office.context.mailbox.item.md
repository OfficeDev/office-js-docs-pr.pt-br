---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,7
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 87ac31deaabda5fe1d2ca024972e326cfa168ea7
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068249"
---
# <a name="item"></a><span data-ttu-id="da986-102">item</span><span class="sxs-lookup"><span data-stu-id="da986-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="da986-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="da986-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="da986-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="da986-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-106">Requirements</span></span>

|<span data-ttu-id="da986-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-107">Requirement</span></span>|<span data-ttu-id="da986-108">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-110">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-110">1.0</span></span>|
|[<span data-ttu-id="da986-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="da986-112">Restricted</span></span>|
|[<span data-ttu-id="da986-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-114">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="da986-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="da986-115">Members and methods</span></span>

| <span data-ttu-id="da986-116">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-116">Member</span></span> | <span data-ttu-id="da986-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="da986-118">attachments</span><span class="sxs-lookup"><span data-stu-id="da986-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="da986-119">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-119">Member</span></span> |
| [<span data-ttu-id="da986-120">bcc</span><span class="sxs-lookup"><span data-stu-id="da986-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="da986-121">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-121">Member</span></span> |
| [<span data-ttu-id="da986-122">body</span><span class="sxs-lookup"><span data-stu-id="da986-122">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="da986-123">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-123">Member</span></span> |
| [<span data-ttu-id="da986-124">cc</span><span class="sxs-lookup"><span data-stu-id="da986-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="da986-125">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-125">Member</span></span> |
| [<span data-ttu-id="da986-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="da986-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="da986-127">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-127">Member</span></span> |
| [<span data-ttu-id="da986-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="da986-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="da986-129">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-129">Member</span></span> |
| [<span data-ttu-id="da986-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="da986-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="da986-131">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-131">Member</span></span> |
| [<span data-ttu-id="da986-132">end</span><span class="sxs-lookup"><span data-stu-id="da986-132">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="da986-133">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-133">Member</span></span> |
| [<span data-ttu-id="da986-134">from</span><span class="sxs-lookup"><span data-stu-id="da986-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="da986-135">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-135">Member</span></span> |
| [<span data-ttu-id="da986-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="da986-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="da986-137">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-137">Member</span></span> |
| [<span data-ttu-id="da986-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="da986-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="da986-139">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-139">Member</span></span> |
| [<span data-ttu-id="da986-140">itemId</span><span class="sxs-lookup"><span data-stu-id="da986-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="da986-141">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-141">Member</span></span> |
| [<span data-ttu-id="da986-142">itemType</span><span class="sxs-lookup"><span data-stu-id="da986-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="da986-143">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-143">Member</span></span> |
| [<span data-ttu-id="da986-144">location</span><span class="sxs-lookup"><span data-stu-id="da986-144">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="da986-145">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-145">Member</span></span> |
| [<span data-ttu-id="da986-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="da986-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="da986-147">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-147">Member</span></span> |
| [<span data-ttu-id="da986-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="da986-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="da986-149">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-149">Member</span></span> |
| [<span data-ttu-id="da986-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="da986-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="da986-151">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-151">Member</span></span> |
| [<span data-ttu-id="da986-152">organizer</span><span class="sxs-lookup"><span data-stu-id="da986-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="da986-153">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-153">Member</span></span> |
| [<span data-ttu-id="da986-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="da986-154">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="da986-155">Member</span><span class="sxs-lookup"><span data-stu-id="da986-155">Member</span></span> |
| [<span data-ttu-id="da986-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="da986-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="da986-157">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-157">Member</span></span> |
| [<span data-ttu-id="da986-158">sender</span><span class="sxs-lookup"><span data-stu-id="da986-158">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="da986-159">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-159">Member</span></span> |
| [<span data-ttu-id="da986-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="da986-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="da986-161">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-161">Member</span></span> |
| [<span data-ttu-id="da986-162">start</span><span class="sxs-lookup"><span data-stu-id="da986-162">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="da986-163">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-163">Member</span></span> |
| [<span data-ttu-id="da986-164">subject</span><span class="sxs-lookup"><span data-stu-id="da986-164">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="da986-165">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-165">Member</span></span> |
| [<span data-ttu-id="da986-166">to</span><span class="sxs-lookup"><span data-stu-id="da986-166">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="da986-167">Membro</span><span class="sxs-lookup"><span data-stu-id="da986-167">Member</span></span> |
| [<span data-ttu-id="da986-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="da986-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="da986-169">Método</span><span class="sxs-lookup"><span data-stu-id="da986-169">Method</span></span> |
| [<span data-ttu-id="da986-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="da986-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="da986-171">Método</span><span class="sxs-lookup"><span data-stu-id="da986-171">Method</span></span> |
| [<span data-ttu-id="da986-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="da986-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="da986-173">Método</span><span class="sxs-lookup"><span data-stu-id="da986-173">Method</span></span> |
| [<span data-ttu-id="da986-174">close</span><span class="sxs-lookup"><span data-stu-id="da986-174">close</span></span>](#close) | <span data-ttu-id="da986-175">Método</span><span class="sxs-lookup"><span data-stu-id="da986-175">Method</span></span> |
| [<span data-ttu-id="da986-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="da986-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="da986-177">Método</span><span class="sxs-lookup"><span data-stu-id="da986-177">Method</span></span> |
| [<span data-ttu-id="da986-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="da986-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="da986-179">Método</span><span class="sxs-lookup"><span data-stu-id="da986-179">Method</span></span> |
| [<span data-ttu-id="da986-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="da986-180">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="da986-181">Método</span><span class="sxs-lookup"><span data-stu-id="da986-181">Method</span></span> |
| [<span data-ttu-id="da986-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="da986-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="da986-183">Método</span><span class="sxs-lookup"><span data-stu-id="da986-183">Method</span></span> |
| [<span data-ttu-id="da986-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="da986-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="da986-185">Método</span><span class="sxs-lookup"><span data-stu-id="da986-185">Method</span></span> |
| [<span data-ttu-id="da986-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="da986-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="da986-187">Método</span><span class="sxs-lookup"><span data-stu-id="da986-187">Method</span></span> |
| [<span data-ttu-id="da986-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="da986-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="da986-189">Método</span><span class="sxs-lookup"><span data-stu-id="da986-189">Method</span></span> |
| [<span data-ttu-id="da986-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="da986-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="da986-191">Método</span><span class="sxs-lookup"><span data-stu-id="da986-191">Method</span></span> |
| [<span data-ttu-id="da986-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="da986-192">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="da986-193">Método</span><span class="sxs-lookup"><span data-stu-id="da986-193">Method</span></span> |
| [<span data-ttu-id="da986-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="da986-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="da986-195">Método</span><span class="sxs-lookup"><span data-stu-id="da986-195">Method</span></span> |
| [<span data-ttu-id="da986-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="da986-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="da986-197">Método</span><span class="sxs-lookup"><span data-stu-id="da986-197">Method</span></span> |
| [<span data-ttu-id="da986-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="da986-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="da986-199">Método</span><span class="sxs-lookup"><span data-stu-id="da986-199">Method</span></span> |
| [<span data-ttu-id="da986-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="da986-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="da986-201">Método</span><span class="sxs-lookup"><span data-stu-id="da986-201">Method</span></span> |
| [<span data-ttu-id="da986-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="da986-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="da986-203">Método</span><span class="sxs-lookup"><span data-stu-id="da986-203">Method</span></span> |
| [<span data-ttu-id="da986-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="da986-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="da986-205">Método</span><span class="sxs-lookup"><span data-stu-id="da986-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="da986-206">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-206">Example</span></span>

<span data-ttu-id="da986-207">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="da986-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="da986-208">Membros</span><span class="sxs-lookup"><span data-stu-id="da986-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="da986-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="da986-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="da986-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="da986-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-212">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="da986-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="da986-213">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="da986-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="da986-214">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-214">Type</span></span>

*   <span data-ttu-id="da986-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="da986-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-216">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-216">Requirements</span></span>

|<span data-ttu-id="da986-217">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-217">Requirement</span></span>|<span data-ttu-id="da986-218">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-219">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-220">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-220">1.0</span></span>|
|[<span data-ttu-id="da986-221">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-222">ReadItem</span></span>|
|[<span data-ttu-id="da986-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-224">Read</span><span class="sxs-lookup"><span data-stu-id="da986-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-225">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-225">Example</span></span>

<span data-ttu-id="da986-226">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="da986-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="da986-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="da986-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="da986-228">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="da986-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="da986-229">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="da986-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="da986-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-230">Type</span></span>

*   [<span data-ttu-id="da986-231">Destinatários</span><span class="sxs-lookup"><span data-stu-id="da986-231">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="da986-232">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-232">Requirements</span></span>

|<span data-ttu-id="da986-233">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-233">Requirement</span></span>|<span data-ttu-id="da986-234">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-235">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-236">1.1</span><span class="sxs-lookup"><span data-stu-id="da986-236">1.1</span></span>|
|[<span data-ttu-id="da986-237">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-237">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-238">ReadItem</span></span>|
|[<span data-ttu-id="da986-239">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-239">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-240">Escrever</span><span class="sxs-lookup"><span data-stu-id="da986-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-241">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-241">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="da986-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="da986-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="da986-243">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="da986-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="da986-244">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-244">Type</span></span>

*   [<span data-ttu-id="da986-245">Corpo</span><span class="sxs-lookup"><span data-stu-id="da986-245">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="da986-246">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-246">Requirements</span></span>

|<span data-ttu-id="da986-247">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-247">Requirement</span></span>|<span data-ttu-id="da986-248">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-249">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-250">1.1</span><span class="sxs-lookup"><span data-stu-id="da986-250">1.1</span></span>|
|[<span data-ttu-id="da986-251">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-252">ReadItem</span></span>|
|[<span data-ttu-id="da986-253">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-254">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-254">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-255">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-255">Example</span></span>

<span data-ttu-id="da986-256">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="da986-256">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="da986-257">A seguir apresentamos um exemplo do parâmetro result passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-257">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="da986-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="da986-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="da986-259">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="da986-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="da986-260">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="da986-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da986-261">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="da986-261">Read mode</span></span>

<span data-ttu-id="da986-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="da986-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="da986-264">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="da986-264">Compose mode</span></span>

<span data-ttu-id="da986-265">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="da986-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="da986-266">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-266">Type</span></span>

*   <span data-ttu-id="da986-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="da986-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-268">Requirements</span></span>

|<span data-ttu-id="da986-269">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-269">Requirement</span></span>|<span data-ttu-id="da986-270">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-271">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-272">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-272">1.0</span></span>|
|[<span data-ttu-id="da986-273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-274">ReadItem</span></span>|
|[<span data-ttu-id="da986-275">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-276">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-276">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="da986-277">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="da986-277">(nullable) conversationId :String</span></span>

<span data-ttu-id="da986-278">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="da986-278">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="da986-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="da986-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="da986-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="da986-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="da986-283">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-283">Type</span></span>

*   <span data-ttu-id="da986-284">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="da986-284">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-285">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-285">Requirements</span></span>

|<span data-ttu-id="da986-286">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-286">Requirement</span></span>|<span data-ttu-id="da986-287">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-288">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-289">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-289">1.0</span></span>|
|[<span data-ttu-id="da986-290">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-291">ReadItem</span></span>|
|[<span data-ttu-id="da986-292">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-293">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-294">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-294">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="da986-295">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="da986-295">dateTimeCreated :Date</span></span>

<span data-ttu-id="da986-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="da986-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="da986-298">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-298">Type</span></span>

*   <span data-ttu-id="da986-299">Data</span><span class="sxs-lookup"><span data-stu-id="da986-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-300">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-300">Requirements</span></span>

|<span data-ttu-id="da986-301">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-301">Requirement</span></span>|<span data-ttu-id="da986-302">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-303">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-304">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-304">1.0</span></span>|
|[<span data-ttu-id="da986-305">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-306">ReadItem</span></span>|
|[<span data-ttu-id="da986-307">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-308">Leitura</span><span class="sxs-lookup"><span data-stu-id="da986-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-309">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-309">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="da986-310">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="da986-310">dateTimeModified :Date</span></span>

<span data-ttu-id="da986-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="da986-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-313">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="da986-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="da986-314">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-314">Type</span></span>

*   <span data-ttu-id="da986-315">Data</span><span class="sxs-lookup"><span data-stu-id="da986-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-316">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-316">Requirements</span></span>

|<span data-ttu-id="da986-317">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-317">Requirement</span></span>|<span data-ttu-id="da986-318">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-319">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-320">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-320">1.0</span></span>|
|[<span data-ttu-id="da986-321">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-322">ReadItem</span></span>|
|[<span data-ttu-id="da986-323">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-324">Leitura</span><span class="sxs-lookup"><span data-stu-id="da986-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-325">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-325">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="da986-326">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="da986-326">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="da986-327">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="da986-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="da986-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="da986-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da986-330">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="da986-330">Read mode</span></span>

<span data-ttu-id="da986-331">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="da986-331">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="da986-332">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="da986-332">Compose mode</span></span>

<span data-ttu-id="da986-333">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="da986-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="da986-334">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="da986-334">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="da986-335">O exemplo a seguir define a hora de término de um compromisso usando [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) o método do `Time` objeto.</span><span class="sxs-lookup"><span data-stu-id="da986-335">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="da986-336">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-336">Type</span></span>

*   <span data-ttu-id="da986-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="da986-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-338">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-338">Requirements</span></span>

|<span data-ttu-id="da986-339">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-339">Requirement</span></span>|<span data-ttu-id="da986-340">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-341">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-342">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-342">1.0</span></span>|
|[<span data-ttu-id="da986-343">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-344">ReadItem</span></span>|
|[<span data-ttu-id="da986-345">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-346">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-346">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="da986-347">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="da986-347">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="da986-348">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="da986-348">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="da986-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="da986-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-351">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="da986-351">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da986-352">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="da986-352">Read mode</span></span>

<span data-ttu-id="da986-353">A propriedade `from` retorna um objeto `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="da986-353">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="da986-354">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="da986-354">Compose mode</span></span>

<span data-ttu-id="da986-355">A propriedade `from` retorna um objeto `From` que fornece um método para obtenção do valor de from.</span><span class="sxs-lookup"><span data-stu-id="da986-355">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="da986-356">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-356">Type</span></span>

*   <span data-ttu-id="da986-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="da986-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-358">Requirements</span></span>

|<span data-ttu-id="da986-359">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-359">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="da986-360">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-361">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-361">1.0</span></span>|<span data-ttu-id="da986-362">1.7</span><span class="sxs-lookup"><span data-stu-id="da986-362">1.7</span></span>|
|[<span data-ttu-id="da986-363">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-364">ReadItem</span></span>|<span data-ttu-id="da986-365">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da986-365">ReadWriteItem</span></span>|
|[<span data-ttu-id="da986-366">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-367">Read</span><span class="sxs-lookup"><span data-stu-id="da986-367">Read</span></span>|<span data-ttu-id="da986-368">Escrever</span><span class="sxs-lookup"><span data-stu-id="da986-368">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="da986-369">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="da986-369">internetMessageId :String</span></span>

<span data-ttu-id="da986-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="da986-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="da986-372">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-372">Type</span></span>

*   <span data-ttu-id="da986-373">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="da986-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-374">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-374">Requirements</span></span>

|<span data-ttu-id="da986-375">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-375">Requirement</span></span>|<span data-ttu-id="da986-376">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-377">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-378">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-378">1.0</span></span>|
|[<span data-ttu-id="da986-379">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-379">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-380">ReadItem</span></span>|
|[<span data-ttu-id="da986-381">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-381">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-382">Leitura</span><span class="sxs-lookup"><span data-stu-id="da986-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-383">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-383">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="da986-384">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="da986-384">itemClass :String</span></span>

<span data-ttu-id="da986-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="da986-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="da986-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="da986-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="da986-389">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-389">Type</span></span>|<span data-ttu-id="da986-390">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-390">Description</span></span>|<span data-ttu-id="da986-391">classe de item</span><span class="sxs-lookup"><span data-stu-id="da986-391">item class</span></span>|
|---|---|---|
|<span data-ttu-id="da986-392">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="da986-392">Appointment items</span></span>|<span data-ttu-id="da986-393">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="da986-393">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="da986-394">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="da986-394">Message items</span></span>|<span data-ttu-id="da986-395">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="da986-395">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="da986-396">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="da986-396">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="da986-397">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-397">Type</span></span>

*   <span data-ttu-id="da986-398">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="da986-398">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-399">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-399">Requirements</span></span>

|<span data-ttu-id="da986-400">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-400">Requirement</span></span>|<span data-ttu-id="da986-401">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-402">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-403">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-403">1.0</span></span>|
|[<span data-ttu-id="da986-404">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-404">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-405">ReadItem</span></span>|
|[<span data-ttu-id="da986-406">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-406">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-407">Read</span><span class="sxs-lookup"><span data-stu-id="da986-407">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-408">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-408">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="da986-409">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="da986-409">(nullable) itemId :String</span></span>

<span data-ttu-id="da986-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="da986-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-412">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="da986-412">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="da986-413">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="da986-413">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="da986-414">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="da986-414">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="da986-415">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="da986-415">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="da986-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="da986-418">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-418">Type</span></span>

*   <span data-ttu-id="da986-419">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="da986-419">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-420">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-420">Requirements</span></span>

|<span data-ttu-id="da986-421">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-421">Requirement</span></span>|<span data-ttu-id="da986-422">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-423">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-424">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-424">1.0</span></span>|
|[<span data-ttu-id="da986-425">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-426">ReadItem</span></span>|
|[<span data-ttu-id="da986-427">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-428">Read</span><span class="sxs-lookup"><span data-stu-id="da986-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-429">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-429">Example</span></span>

<span data-ttu-id="da986-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="da986-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="da986-432">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="da986-432">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="da986-433">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="da986-433">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="da986-434">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="da986-434">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="da986-435">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-435">Type</span></span>

*   [<span data-ttu-id="da986-436">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="da986-436">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="da986-437">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-437">Requirements</span></span>

|<span data-ttu-id="da986-438">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-438">Requirement</span></span>|<span data-ttu-id="da986-439">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-440">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-441">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-441">1.0</span></span>|
|[<span data-ttu-id="da986-442">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-442">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-443">ReadItem</span></span>|
|[<span data-ttu-id="da986-444">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-444">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-445">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-445">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-446">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-446">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="da986-447">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="da986-447">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="da986-448">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="da986-448">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da986-449">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="da986-449">Read mode</span></span>

<span data-ttu-id="da986-450">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="da986-450">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="da986-451">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="da986-451">Compose mode</span></span>

<span data-ttu-id="da986-452">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="da986-452">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="da986-453">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-453">Type</span></span>

*   <span data-ttu-id="da986-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="da986-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-455">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-455">Requirements</span></span>

|<span data-ttu-id="da986-456">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-456">Requirement</span></span>|<span data-ttu-id="da986-457">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-458">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-459">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-459">1.0</span></span>|
|[<span data-ttu-id="da986-460">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-460">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-461">ReadItem</span></span>|
|[<span data-ttu-id="da986-462">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-462">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-463">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-463">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="da986-464">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="da986-464">normalizedSubject :String</span></span>

<span data-ttu-id="da986-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="da986-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="da986-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject).</span><span class="sxs-lookup"><span data-stu-id="da986-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="da986-469">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-469">Type</span></span>

*   <span data-ttu-id="da986-470">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="da986-470">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-471">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-471">Requirements</span></span>

|<span data-ttu-id="da986-472">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-472">Requirement</span></span>|<span data-ttu-id="da986-473">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-474">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-475">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-475">1.0</span></span>|
|[<span data-ttu-id="da986-476">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-476">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-477">ReadItem</span></span>|
|[<span data-ttu-id="da986-478">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-478">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-479">Read</span><span class="sxs-lookup"><span data-stu-id="da986-479">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-480">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-480">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="da986-481">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="da986-481">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="da986-482">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="da986-482">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="da986-483">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-483">Type</span></span>

*   [<span data-ttu-id="da986-484">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="da986-484">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="da986-485">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-485">Requirements</span></span>

|<span data-ttu-id="da986-486">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-486">Requirement</span></span>|<span data-ttu-id="da986-487">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-488">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-489">1.3</span><span class="sxs-lookup"><span data-stu-id="da986-489">1.3</span></span>|
|[<span data-ttu-id="da986-490">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-491">ReadItem</span></span>|
|[<span data-ttu-id="da986-492">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-493">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-493">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-494">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-494">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="da986-495">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="da986-495">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="da986-496">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="da986-496">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="da986-497">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="da986-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da986-498">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="da986-498">Read mode</span></span>

<span data-ttu-id="da986-499">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="da986-499">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="da986-500">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="da986-500">Compose mode</span></span>

<span data-ttu-id="da986-501">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="da986-501">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="da986-502">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-502">Type</span></span>

*   <span data-ttu-id="da986-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="da986-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-504">Requirements</span></span>

|<span data-ttu-id="da986-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-505">Requirement</span></span>|<span data-ttu-id="da986-506">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-507">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-508">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-508">1.0</span></span>|
|[<span data-ttu-id="da986-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-510">ReadItem</span></span>|
|[<span data-ttu-id="da986-511">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-512">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-512">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="da986-513">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="da986-513">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="da986-514">Obtém o endereço de email do organizador para uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="da986-514">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da986-515">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="da986-515">Read mode</span></span>

<span data-ttu-id="da986-516">A propriedade `organizer` retorna um objeto [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="da986-516">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="da986-517">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="da986-517">Compose mode</span></span>

<span data-ttu-id="da986-518">A propriedade `organizer` retorna um objeto [Organizer](/javascript/api/outlook_1_7/office.organizer) que fornece um método para obtenção do valor de organizer.</span><span class="sxs-lookup"><span data-stu-id="da986-518">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="da986-519">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-519">Type</span></span>

*   <span data-ttu-id="da986-520">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="da986-520">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-521">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-521">Requirements</span></span>

|<span data-ttu-id="da986-522">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-522">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="da986-523">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-524">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-524">1.0</span></span>|<span data-ttu-id="da986-525">1.7</span><span class="sxs-lookup"><span data-stu-id="da986-525">1.7</span></span>|
|[<span data-ttu-id="da986-526">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-526">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-527">ReadItem</span></span>|<span data-ttu-id="da986-528">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da986-528">ReadWriteItem</span></span>|
|[<span data-ttu-id="da986-529">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-530">Read</span><span class="sxs-lookup"><span data-stu-id="da986-530">Read</span></span>|<span data-ttu-id="da986-531">Escrever</span><span class="sxs-lookup"><span data-stu-id="da986-531">Compose</span></span>|

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="da986-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="da986-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="da986-533">Obtém ou configura o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="da986-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="da986-534">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="da986-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="da986-535">Modos de leitura e redação para itens do compromisso.</span><span class="sxs-lookup"><span data-stu-id="da986-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="da986-536">Modo de leitura para os itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="da986-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="da986-537">A propriedade `recurrence` retorna um objeto [recurrence](/javascript/api/outlook_1_7/office.recurrence) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="da986-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="da986-538">`null` retorna para compromissos individuais e solicitações de reunião de compromissos individuais.</span><span class="sxs-lookup"><span data-stu-id="da986-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="da986-539">`undefined` retorna para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="da986-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="da986-540">Observação: solicitações de reunião têm um valor `itemClass` de IPM.Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="da986-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="da986-541">Observação: se o objeto de recorrência for `null`, isso indicará que o objeto é um compromisso individual ou uma solicitação de reunião de um compromisso individual e NÃO parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="da986-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da986-542">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="da986-542">Read mode</span></span>

<span data-ttu-id="da986-543">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook_1_7/office.recurrence) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="da986-543">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="da986-544">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="da986-544">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="da986-545">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="da986-545">Compose mode</span></span>

<span data-ttu-id="da986-546">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook_1_7/office.recurrence) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="da986-546">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="da986-547">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="da986-547">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="da986-548">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-548">Type</span></span>

* [<span data-ttu-id="da986-549">Recurrence</span><span class="sxs-lookup"><span data-stu-id="da986-549">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="da986-550">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-550">Requirement</span></span>|<span data-ttu-id="da986-551">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-551">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-552">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-552">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-553">1.7</span><span class="sxs-lookup"><span data-stu-id="da986-553">1.7</span></span>|
|[<span data-ttu-id="da986-554">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-554">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-555">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-555">ReadItem</span></span>|
|[<span data-ttu-id="da986-556">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-556">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-557">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-557">Compose or Read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="da986-558">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="da986-558">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="da986-559">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="da986-559">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="da986-560">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="da986-560">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da986-561">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="da986-561">Read mode</span></span>

<span data-ttu-id="da986-562">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="da986-562">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="da986-563">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="da986-563">Compose mode</span></span>

<span data-ttu-id="da986-564">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="da986-564">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="da986-565">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-565">Type</span></span>

*   <span data-ttu-id="da986-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="da986-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-567">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-567">Requirements</span></span>

|<span data-ttu-id="da986-568">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-568">Requirement</span></span>|<span data-ttu-id="da986-569">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-570">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-571">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-571">1.0</span></span>|
|[<span data-ttu-id="da986-572">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-572">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-573">ReadItem</span></span>|
|[<span data-ttu-id="da986-574">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-575">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-575">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="da986-576">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="da986-576">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="da986-p128">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="da986-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="da986-p129">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="da986-p129">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-581">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="da986-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="da986-582">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-582">Type</span></span>

*   [<span data-ttu-id="da986-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="da986-583">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="da986-584">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-584">Requirements</span></span>

|<span data-ttu-id="da986-585">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-585">Requirement</span></span>|<span data-ttu-id="da986-586">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-587">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-588">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-588">1.0</span></span>|
|[<span data-ttu-id="da986-589">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-590">ReadItem</span></span>|
|[<span data-ttu-id="da986-591">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-592">Leitura</span><span class="sxs-lookup"><span data-stu-id="da986-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-593">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-593">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="da986-594">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="da986-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="da986-595">Obtém a id da série a qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="da986-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="da986-596">No OWA e no Outlook, o `seriesId` retorna a ID dos Serviços Web do Exchange (EWS) do item pai (série) a qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="da986-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="da986-597">No entanto, no iOS e no Android, o `seriesId` retorna a ID REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="da986-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-598">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="da986-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="da986-599">A propriedade `seriesId` não é idêntica à ID do Outlook usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="da986-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="da986-600">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="da986-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="da986-601">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="da986-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="da986-602">A propriedade `seriesId` retorna `null` para itens que não têm itens pai como compromissos individuais, itens de série ou solicitações de reunião e retorna `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="da986-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="da986-603">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-603">Type</span></span>

* <span data-ttu-id="da986-604">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="da986-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-605">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-605">Requirements</span></span>

|<span data-ttu-id="da986-606">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-606">Requirement</span></span>|<span data-ttu-id="da986-607">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-608">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-609">1.7</span><span class="sxs-lookup"><span data-stu-id="da986-609">1.7</span></span>|
|[<span data-ttu-id="da986-610">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-611">ReadItem</span></span>|
|[<span data-ttu-id="da986-612">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-613">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-613">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-614">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-614">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="da986-615">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="da986-615">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="da986-616">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="da986-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="da986-p132">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="da986-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da986-619">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="da986-619">Read mode</span></span>

<span data-ttu-id="da986-620">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="da986-620">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="da986-621">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="da986-621">Compose mode</span></span>

<span data-ttu-id="da986-622">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="da986-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="da986-623">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="da986-623">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="da986-624">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="da986-624">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="da986-625">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-625">Type</span></span>

*   <span data-ttu-id="da986-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="da986-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-627">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-627">Requirements</span></span>

|<span data-ttu-id="da986-628">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-628">Requirement</span></span>|<span data-ttu-id="da986-629">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-629">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-630">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-630">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-631">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-631">1.0</span></span>|
|[<span data-ttu-id="da986-632">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-632">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-633">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-633">ReadItem</span></span>|
|[<span data-ttu-id="da986-634">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-634">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-635">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-635">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="da986-636">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="da986-636">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="da986-637">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="da986-637">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="da986-638">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="da986-638">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da986-639">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="da986-639">Read mode</span></span>

<span data-ttu-id="da986-p133">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="da986-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="da986-642">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="da986-642">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="da986-643">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="da986-643">Compose mode</span></span>

<span data-ttu-id="da986-644">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="da986-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="da986-645">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-645">Type</span></span>

*   <span data-ttu-id="da986-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="da986-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-647">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-647">Requirements</span></span>

|<span data-ttu-id="da986-648">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-648">Requirement</span></span>|<span data-ttu-id="da986-649">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-650">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-651">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-651">1.0</span></span>|
|[<span data-ttu-id="da986-652">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-653">ReadItem</span></span>|
|[<span data-ttu-id="da986-654">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-655">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-655">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="da986-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="da986-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="da986-657">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="da986-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="da986-658">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="da986-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="da986-659">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="da986-659">Read mode</span></span>

<span data-ttu-id="da986-p135">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="da986-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="da986-662">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="da986-662">Compose mode</span></span>

<span data-ttu-id="da986-663">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="da986-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="da986-664">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-664">Type</span></span>

*   <span data-ttu-id="da986-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="da986-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-666">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-666">Requirements</span></span>

|<span data-ttu-id="da986-667">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-667">Requirement</span></span>|<span data-ttu-id="da986-668">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-669">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-670">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-670">1.0</span></span>|
|[<span data-ttu-id="da986-671">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-672">ReadItem</span></span>|
|[<span data-ttu-id="da986-673">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-674">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-674">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="da986-675">Métodos</span><span class="sxs-lookup"><span data-stu-id="da986-675">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="da986-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="da986-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="da986-677">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="da986-677">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="da986-678">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="da986-678">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="da986-679">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="da986-679">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-680">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-680">Parameters</span></span>
|<span data-ttu-id="da986-681">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-681">Name</span></span>|<span data-ttu-id="da986-682">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-682">Type</span></span>|<span data-ttu-id="da986-683">Atributos</span><span class="sxs-lookup"><span data-stu-id="da986-683">Attributes</span></span>|<span data-ttu-id="da986-684">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-684">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="da986-685">String</span><span class="sxs-lookup"><span data-stu-id="da986-685">String</span></span>||<span data-ttu-id="da986-p136">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="da986-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="da986-688">String</span><span class="sxs-lookup"><span data-stu-id="da986-688">String</span></span>||<span data-ttu-id="da986-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="da986-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="da986-691">Object</span><span class="sxs-lookup"><span data-stu-id="da986-691">Object</span></span>|<span data-ttu-id="da986-692">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-692">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-693">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="da986-693">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="da986-694">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-694">Object</span></span>|<span data-ttu-id="da986-695">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-695">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-696">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-696">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="da986-697">Booliano</span><span class="sxs-lookup"><span data-stu-id="da986-697">Boolean</span></span>|<span data-ttu-id="da986-698">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-698">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-699">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="da986-699">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="da986-700">function</span><span class="sxs-lookup"><span data-stu-id="da986-700">function</span></span>|<span data-ttu-id="da986-701">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-701">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-702">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da986-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="da986-703">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="da986-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="da986-704">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="da986-704">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="da986-705">Erros</span><span class="sxs-lookup"><span data-stu-id="da986-705">Errors</span></span>

|<span data-ttu-id="da986-706">Código de erro</span><span class="sxs-lookup"><span data-stu-id="da986-706">Error code</span></span>|<span data-ttu-id="da986-707">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-707">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="da986-708">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="da986-708">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="da986-709">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="da986-709">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="da986-710">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="da986-710">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-711">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-711">Requirements</span></span>

|<span data-ttu-id="da986-712">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-712">Requirement</span></span>|<span data-ttu-id="da986-713">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-713">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-714">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-714">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-715">1.1</span><span class="sxs-lookup"><span data-stu-id="da986-715">1.1</span></span>|
|[<span data-ttu-id="da986-716">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-716">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-717">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da986-717">ReadWriteItem</span></span>|
|[<span data-ttu-id="da986-718">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-718">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-719">Escrever</span><span class="sxs-lookup"><span data-stu-id="da986-719">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="da986-720">Exemplos</span><span class="sxs-lookup"><span data-stu-id="da986-720">Examples</span></span>

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

<span data-ttu-id="da986-721">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="da986-721">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="da986-722">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="da986-722">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="da986-723">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="da986-723">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="da986-724">Atualmente, os tipos de evento compatíveis são `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="da986-724">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-725">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-725">Parameters</span></span>

| <span data-ttu-id="da986-726">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-726">Name</span></span> | <span data-ttu-id="da986-727">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-727">Type</span></span> | <span data-ttu-id="da986-728">Atributos</span><span class="sxs-lookup"><span data-stu-id="da986-728">Attributes</span></span> | <span data-ttu-id="da986-729">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-729">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="da986-730">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="da986-730">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="da986-731">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="da986-731">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="da986-732">Função</span><span class="sxs-lookup"><span data-stu-id="da986-732">Function</span></span> || <span data-ttu-id="da986-p138">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="da986-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="da986-736">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-736">Object</span></span> | <span data-ttu-id="da986-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-737">&lt;optional&gt;</span></span> | <span data-ttu-id="da986-738">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="da986-738">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="da986-739">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-739">Object</span></span> | <span data-ttu-id="da986-740">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-740">&lt;optional&gt;</span></span> | <span data-ttu-id="da986-741">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-741">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="da986-742">function</span><span class="sxs-lookup"><span data-stu-id="da986-742">function</span></span>| <span data-ttu-id="da986-743">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-743">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-744">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da986-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-745">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-745">Requirements</span></span>

|<span data-ttu-id="da986-746">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-746">Requirement</span></span>| <span data-ttu-id="da986-747">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-748">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da986-749">1.7</span><span class="sxs-lookup"><span data-stu-id="da986-749">1.7</span></span> |
|[<span data-ttu-id="da986-750">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da986-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-751">ReadItem</span></span> |
|[<span data-ttu-id="da986-752">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="da986-753">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-753">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="da986-754">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-754">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="da986-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="da986-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="da986-756">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="da986-756">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="da986-p139">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="da986-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="da986-760">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="da986-760">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="da986-761">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="da986-761">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-762">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-762">Parameters</span></span>

|<span data-ttu-id="da986-763">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-763">Name</span></span>|<span data-ttu-id="da986-764">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-764">Type</span></span>|<span data-ttu-id="da986-765">Atributos</span><span class="sxs-lookup"><span data-stu-id="da986-765">Attributes</span></span>|<span data-ttu-id="da986-766">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-766">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="da986-767">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="da986-767">String</span></span>||<span data-ttu-id="da986-p140">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="da986-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="da986-770">String</span><span class="sxs-lookup"><span data-stu-id="da986-770">String</span></span>||<span data-ttu-id="da986-771">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="da986-771">The subject of the item to be attached.</span></span> <span data-ttu-id="da986-772">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="da986-772">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="da986-773">Object</span><span class="sxs-lookup"><span data-stu-id="da986-773">Object</span></span>|<span data-ttu-id="da986-774">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-774">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-775">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="da986-775">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="da986-776">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-776">Object</span></span>|<span data-ttu-id="da986-777">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-777">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-778">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-778">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="da986-779">function</span><span class="sxs-lookup"><span data-stu-id="da986-779">function</span></span>|<span data-ttu-id="da986-780">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-780">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-781">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da986-781">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="da986-782">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="da986-782">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="da986-783">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="da986-783">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="da986-784">Erros</span><span class="sxs-lookup"><span data-stu-id="da986-784">Errors</span></span>

|<span data-ttu-id="da986-785">Código de erro</span><span class="sxs-lookup"><span data-stu-id="da986-785">Error code</span></span>|<span data-ttu-id="da986-786">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-786">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="da986-787">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="da986-787">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-788">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-788">Requirements</span></span>

|<span data-ttu-id="da986-789">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-789">Requirement</span></span>|<span data-ttu-id="da986-790">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-791">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-792">1.1</span><span class="sxs-lookup"><span data-stu-id="da986-792">1.1</span></span>|
|[<span data-ttu-id="da986-793">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-793">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-794">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da986-794">ReadWriteItem</span></span>|
|[<span data-ttu-id="da986-795">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-795">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-796">Escrever</span><span class="sxs-lookup"><span data-stu-id="da986-796">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-797">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-797">Example</span></span>

<span data-ttu-id="da986-798">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="da986-798">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="da986-799">close()</span><span class="sxs-lookup"><span data-stu-id="da986-799">close()</span></span>

<span data-ttu-id="da986-800">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="da986-800">Closes the current item that is being composed.</span></span>

<span data-ttu-id="da986-p142">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="da986-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-803">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="da986-803">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="da986-804">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="da986-804">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-805">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-805">Requirements</span></span>

|<span data-ttu-id="da986-806">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-806">Requirement</span></span>|<span data-ttu-id="da986-807">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-808">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-809">1.3</span><span class="sxs-lookup"><span data-stu-id="da986-809">1.3</span></span>|
|[<span data-ttu-id="da986-810">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-811">Restrito</span><span class="sxs-lookup"><span data-stu-id="da986-811">Restricted</span></span>|
|[<span data-ttu-id="da986-812">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-813">Escrever</span><span class="sxs-lookup"><span data-stu-id="da986-813">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="da986-814">displayReplyAllForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="da986-814">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="da986-815">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="da986-815">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-816">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="da986-816">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="da986-817">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="da986-817">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="da986-818">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="da986-818">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="da986-p143">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="da986-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-822">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-822">Parameters</span></span>

|<span data-ttu-id="da986-823">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-823">Name</span></span>|<span data-ttu-id="da986-824">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-824">Type</span></span>|<span data-ttu-id="da986-825">Atributos</span><span class="sxs-lookup"><span data-stu-id="da986-825">Attributes</span></span>|<span data-ttu-id="da986-826">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-826">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="da986-827">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="da986-827">String &#124; Object</span></span>||<span data-ttu-id="da986-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="da986-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="da986-830">**OU**</span><span class="sxs-lookup"><span data-stu-id="da986-830">**OR**</span></span><br/><span data-ttu-id="da986-p145">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="da986-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="da986-833">String</span><span class="sxs-lookup"><span data-stu-id="da986-833">String</span></span>|<span data-ttu-id="da986-834">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-834">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="da986-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="da986-837">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-837">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="da986-838">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-838">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-839">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="da986-839">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="da986-840">String</span><span class="sxs-lookup"><span data-stu-id="da986-840">String</span></span>||<span data-ttu-id="da986-p147">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="da986-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="da986-843">String</span><span class="sxs-lookup"><span data-stu-id="da986-843">String</span></span>||<span data-ttu-id="da986-844">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="da986-844">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="da986-845">String</span><span class="sxs-lookup"><span data-stu-id="da986-845">String</span></span>||<span data-ttu-id="da986-p148">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="da986-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="da986-848">Booliano</span><span class="sxs-lookup"><span data-stu-id="da986-848">Boolean</span></span>||<span data-ttu-id="da986-p149">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="da986-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="da986-851">String</span><span class="sxs-lookup"><span data-stu-id="da986-851">String</span></span>||<span data-ttu-id="da986-p150">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="da986-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="da986-855">function</span><span class="sxs-lookup"><span data-stu-id="da986-855">function</span></span>|<span data-ttu-id="da986-856">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-856">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-857">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da986-857">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-858">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-858">Requirements</span></span>

|<span data-ttu-id="da986-859">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-859">Requirement</span></span>|<span data-ttu-id="da986-860">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-861">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-862">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-862">1.0</span></span>|
|[<span data-ttu-id="da986-863">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-864">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-864">ReadItem</span></span>|
|[<span data-ttu-id="da986-865">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-866">Read</span><span class="sxs-lookup"><span data-stu-id="da986-866">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="da986-867">Exemplos</span><span class="sxs-lookup"><span data-stu-id="da986-867">Examples</span></span>

<span data-ttu-id="da986-868">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="da986-868">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="da986-869">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="da986-869">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="da986-870">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="da986-870">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="da986-871">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="da986-871">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="da986-872">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="da986-872">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="da986-873">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-873">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="da986-874">displayReplyForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="da986-874">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="da986-875">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="da986-875">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-876">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="da986-876">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="da986-877">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="da986-877">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="da986-878">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="da986-878">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="da986-p151">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="da986-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-882">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-882">Parameters</span></span>

|<span data-ttu-id="da986-883">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-883">Name</span></span>|<span data-ttu-id="da986-884">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-884">Type</span></span>|<span data-ttu-id="da986-885">Atributos</span><span class="sxs-lookup"><span data-stu-id="da986-885">Attributes</span></span>|<span data-ttu-id="da986-886">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-886">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="da986-887">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="da986-887">String &#124; Object</span></span>||<span data-ttu-id="da986-p152">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="da986-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="da986-890">**OU**</span><span class="sxs-lookup"><span data-stu-id="da986-890">**OR**</span></span><br/><span data-ttu-id="da986-p153">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="da986-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="da986-893">String</span><span class="sxs-lookup"><span data-stu-id="da986-893">String</span></span>|<span data-ttu-id="da986-894">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-894">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="da986-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="da986-897">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-897">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="da986-898">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-898">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-899">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="da986-899">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="da986-900">String</span><span class="sxs-lookup"><span data-stu-id="da986-900">String</span></span>||<span data-ttu-id="da986-p155">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="da986-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="da986-903">String</span><span class="sxs-lookup"><span data-stu-id="da986-903">String</span></span>||<span data-ttu-id="da986-904">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="da986-904">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="da986-905">String</span><span class="sxs-lookup"><span data-stu-id="da986-905">String</span></span>||<span data-ttu-id="da986-p156">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="da986-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="da986-908">Booliano</span><span class="sxs-lookup"><span data-stu-id="da986-908">Boolean</span></span>||<span data-ttu-id="da986-p157">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="da986-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="da986-911">String</span><span class="sxs-lookup"><span data-stu-id="da986-911">String</span></span>||<span data-ttu-id="da986-p158">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="da986-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="da986-915">function</span><span class="sxs-lookup"><span data-stu-id="da986-915">function</span></span>|<span data-ttu-id="da986-916">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-916">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-917">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da986-917">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-918">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-918">Requirements</span></span>

|<span data-ttu-id="da986-919">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-919">Requirement</span></span>|<span data-ttu-id="da986-920">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-920">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-921">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-921">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-922">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-922">1.0</span></span>|
|[<span data-ttu-id="da986-923">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-923">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-924">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-924">ReadItem</span></span>|
|[<span data-ttu-id="da986-925">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-925">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-926">Read</span><span class="sxs-lookup"><span data-stu-id="da986-926">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="da986-927">Exemplos</span><span class="sxs-lookup"><span data-stu-id="da986-927">Examples</span></span>

<span data-ttu-id="da986-928">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="da986-928">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="da986-929">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="da986-929">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="da986-930">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="da986-930">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="da986-931">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="da986-931">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="da986-932">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="da986-932">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="da986-933">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-933">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="da986-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="da986-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="da986-935">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="da986-935">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-936">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="da986-936">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-937">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-937">Requirements</span></span>

|<span data-ttu-id="da986-938">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-938">Requirement</span></span>|<span data-ttu-id="da986-939">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-940">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-941">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-941">1.0</span></span>|
|[<span data-ttu-id="da986-942">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-942">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-943">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-943">ReadItem</span></span>|
|[<span data-ttu-id="da986-944">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-944">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-945">Read</span><span class="sxs-lookup"><span data-stu-id="da986-945">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="da986-946">Retorna:</span><span class="sxs-lookup"><span data-stu-id="da986-946">Returns:</span></span>

<span data-ttu-id="da986-947">Tipo: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="da986-947">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="da986-948">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-948">Example</span></span>

<span data-ttu-id="da986-949">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="da986-949">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="da986-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="da986-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="da986-951">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="da986-951">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-952">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="da986-952">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-953">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-953">Parameters</span></span>

|<span data-ttu-id="da986-954">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-954">Name</span></span>|<span data-ttu-id="da986-955">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-955">Type</span></span>|<span data-ttu-id="da986-956">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-956">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="da986-957">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="da986-957">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="da986-958">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="da986-958">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-959">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-959">Requirements</span></span>

|<span data-ttu-id="da986-960">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-960">Requirement</span></span>|<span data-ttu-id="da986-961">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-962">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-963">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-963">1.0</span></span>|
|[<span data-ttu-id="da986-964">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-964">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-965">Restrito</span><span class="sxs-lookup"><span data-stu-id="da986-965">Restricted</span></span>|
|[<span data-ttu-id="da986-966">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-966">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-967">Read</span><span class="sxs-lookup"><span data-stu-id="da986-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="da986-968">Retorna:</span><span class="sxs-lookup"><span data-stu-id="da986-968">Returns:</span></span>

<span data-ttu-id="da986-969">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="da986-969">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="da986-970">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="da986-970">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="da986-971">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="da986-971">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="da986-972">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="da986-972">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="da986-973">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="da986-973">Value of `entityType`</span></span>|<span data-ttu-id="da986-974">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="da986-974">Type of objects in returned array</span></span>|<span data-ttu-id="da986-975">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="da986-975">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="da986-976">String</span><span class="sxs-lookup"><span data-stu-id="da986-976">String</span></span>|<span data-ttu-id="da986-977">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="da986-977">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="da986-978">Contato</span><span class="sxs-lookup"><span data-stu-id="da986-978">Contact</span></span>|<span data-ttu-id="da986-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="da986-979">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="da986-980">String</span><span class="sxs-lookup"><span data-stu-id="da986-980">String</span></span>|<span data-ttu-id="da986-981">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="da986-981">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="da986-982">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="da986-982">MeetingSuggestion</span></span>|<span data-ttu-id="da986-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="da986-983">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="da986-984">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="da986-984">PhoneNumber</span></span>|<span data-ttu-id="da986-985">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="da986-985">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="da986-986">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="da986-986">TaskSuggestion</span></span>|<span data-ttu-id="da986-987">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="da986-987">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="da986-988">String</span><span class="sxs-lookup"><span data-stu-id="da986-988">String</span></span>|<span data-ttu-id="da986-989">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="da986-989">**Restricted**</span></span>|

<span data-ttu-id="da986-990">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="da986-990">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="da986-991">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-991">Example</span></span>

<span data-ttu-id="da986-992">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="da986-992">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="da986-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="da986-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="da986-994">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="da986-994">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-995">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="da986-995">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="da986-996">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="da986-996">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-997">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-997">Parameters</span></span>

|<span data-ttu-id="da986-998">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-998">Name</span></span>|<span data-ttu-id="da986-999">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-999">Type</span></span>|<span data-ttu-id="da986-1000">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-1000">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="da986-1001">String</span><span class="sxs-lookup"><span data-stu-id="da986-1001">String</span></span>|<span data-ttu-id="da986-1002">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="da986-1002">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-1003">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-1003">Requirements</span></span>

|<span data-ttu-id="da986-1004">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-1004">Requirement</span></span>|<span data-ttu-id="da986-1005">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-1006">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-1007">1.0</span></span>|
|[<span data-ttu-id="da986-1008">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-1008">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-1009">ReadItem</span></span>|
|[<span data-ttu-id="da986-1010">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-1010">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-1011">Read</span><span class="sxs-lookup"><span data-stu-id="da986-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="da986-1012">Retorna:</span><span class="sxs-lookup"><span data-stu-id="da986-1012">Returns:</span></span>

<span data-ttu-id="da986-p160">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="da986-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="da986-1015">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="da986-1015">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="da986-1016">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="da986-1016">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="da986-1017">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="da986-1017">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-1018">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="da986-1018">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="da986-p161">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="da986-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="da986-1022">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="da986-1022">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="da986-1023">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="da986-1023">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="da986-p162">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="da986-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-1027">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-1027">Requirements</span></span>

|<span data-ttu-id="da986-1028">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-1028">Requirement</span></span>|<span data-ttu-id="da986-1029">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-1029">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-1030">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-1030">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-1031">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-1031">1.0</span></span>|
|[<span data-ttu-id="da986-1032">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-1032">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-1033">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-1033">ReadItem</span></span>|
|[<span data-ttu-id="da986-1034">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-1034">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-1035">Read</span><span class="sxs-lookup"><span data-stu-id="da986-1035">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="da986-1036">Retorna:</span><span class="sxs-lookup"><span data-stu-id="da986-1036">Returns:</span></span>

<span data-ttu-id="da986-p163">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="da986-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="da986-1039">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="da986-1039">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="da986-1040">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-1040">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="da986-1041">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-1041">Example</span></span>

<span data-ttu-id="da986-1042">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="da986-1042">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="da986-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="da986-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="da986-1044">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="da986-1044">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-1045">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="da986-1045">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="da986-1046">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="da986-1046">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="da986-p164">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="da986-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-1049">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-1049">Parameters</span></span>

|<span data-ttu-id="da986-1050">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-1050">Name</span></span>|<span data-ttu-id="da986-1051">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-1051">Type</span></span>|<span data-ttu-id="da986-1052">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-1052">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="da986-1053">String</span><span class="sxs-lookup"><span data-stu-id="da986-1053">String</span></span>|<span data-ttu-id="da986-1054">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="da986-1054">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-1055">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-1055">Requirements</span></span>

|<span data-ttu-id="da986-1056">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-1056">Requirement</span></span>|<span data-ttu-id="da986-1057">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-1057">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-1058">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-1058">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-1059">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-1059">1.0</span></span>|
|[<span data-ttu-id="da986-1060">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-1060">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-1061">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-1061">ReadItem</span></span>|
|[<span data-ttu-id="da986-1062">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-1062">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-1063">Read</span><span class="sxs-lookup"><span data-stu-id="da986-1063">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="da986-1064">Retorna:</span><span class="sxs-lookup"><span data-stu-id="da986-1064">Returns:</span></span>

<span data-ttu-id="da986-1065">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="da986-1065">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="da986-1066">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="da986-1066">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="da986-1067">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="da986-1067">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="da986-1068">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-1068">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="da986-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="da986-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="da986-1070">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="da986-1070">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="da986-p165">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="da986-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-1073">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-1073">Parameters</span></span>

|<span data-ttu-id="da986-1074">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-1074">Name</span></span>|<span data-ttu-id="da986-1075">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-1075">Type</span></span>|<span data-ttu-id="da986-1076">Atributos</span><span class="sxs-lookup"><span data-stu-id="da986-1076">Attributes</span></span>|<span data-ttu-id="da986-1077">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-1077">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="da986-1078">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="da986-1078">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="da986-p166">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="da986-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="da986-1082">Object</span><span class="sxs-lookup"><span data-stu-id="da986-1082">Object</span></span>|<span data-ttu-id="da986-1083">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-1084">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="da986-1084">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="da986-1085">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-1085">Object</span></span>|<span data-ttu-id="da986-1086">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1086">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-1087">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-1087">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="da986-1088">function</span><span class="sxs-lookup"><span data-stu-id="da986-1088">function</span></span>||<span data-ttu-id="da986-1089">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da986-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="da986-1090">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="da986-1090">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="da986-1091">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="da986-1091">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-1092">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-1092">Requirements</span></span>

|<span data-ttu-id="da986-1093">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-1093">Requirement</span></span>|<span data-ttu-id="da986-1094">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-1094">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-1095">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-1095">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-1096">1.2</span><span class="sxs-lookup"><span data-stu-id="da986-1096">1.2</span></span>|
|[<span data-ttu-id="da986-1097">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-1097">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-1098">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da986-1098">ReadWriteItem</span></span>|
|[<span data-ttu-id="da986-1099">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-1099">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-1100">Escrever</span><span class="sxs-lookup"><span data-stu-id="da986-1100">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="da986-1101">Retorna:</span><span class="sxs-lookup"><span data-stu-id="da986-1101">Returns:</span></span>

<span data-ttu-id="da986-1102">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="da986-1102">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="da986-1103">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="da986-1103">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="da986-1104">String</span><span class="sxs-lookup"><span data-stu-id="da986-1104">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="da986-1105">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-1105">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="da986-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="da986-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="da986-p168">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="da986-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="da986-1109">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="da986-1109">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-1110">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-1110">Requirements</span></span>

|<span data-ttu-id="da986-1111">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-1111">Requirement</span></span>|<span data-ttu-id="da986-1112">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-1113">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-1114">1.6</span><span class="sxs-lookup"><span data-stu-id="da986-1114">1.6</span></span>|
|[<span data-ttu-id="da986-1115">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-1115">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-1116">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-1116">ReadItem</span></span>|
|[<span data-ttu-id="da986-1117">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-1117">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-1118">Read</span><span class="sxs-lookup"><span data-stu-id="da986-1118">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="da986-1119">Retorna:</span><span class="sxs-lookup"><span data-stu-id="da986-1119">Returns:</span></span>

<span data-ttu-id="da986-1120">Tipo: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="da986-1120">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="da986-1121">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-1121">Example</span></span>

<span data-ttu-id="da986-1122">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="da986-1122">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="da986-1123">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="da986-1123">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="da986-p169">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="da986-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="da986-1126">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="da986-1126">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="da986-p170">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="da986-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="da986-1130">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="da986-1130">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="da986-1131">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="da986-1131">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="da986-p171">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="da986-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="da986-1135">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-1135">Requirements</span></span>

|<span data-ttu-id="da986-1136">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-1136">Requirement</span></span>|<span data-ttu-id="da986-1137">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-1138">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-1139">1.6</span><span class="sxs-lookup"><span data-stu-id="da986-1139">1.6</span></span>|
|[<span data-ttu-id="da986-1140">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-1140">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-1141">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-1141">ReadItem</span></span>|
|[<span data-ttu-id="da986-1142">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-1142">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-1143">Read</span><span class="sxs-lookup"><span data-stu-id="da986-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="da986-1144">Retorna:</span><span class="sxs-lookup"><span data-stu-id="da986-1144">Returns:</span></span>

<span data-ttu-id="da986-p172">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="da986-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="da986-1147">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-1147">Example</span></span>

<span data-ttu-id="da986-1148">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="da986-1148">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="da986-1149">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="da986-1149">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="da986-1150">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="da986-1150">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="da986-p173">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="da986-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-1154">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-1154">Parameters</span></span>

|<span data-ttu-id="da986-1155">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-1155">Name</span></span>|<span data-ttu-id="da986-1156">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-1156">Type</span></span>|<span data-ttu-id="da986-1157">Atributos</span><span class="sxs-lookup"><span data-stu-id="da986-1157">Attributes</span></span>|<span data-ttu-id="da986-1158">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-1158">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="da986-1159">function</span><span class="sxs-lookup"><span data-stu-id="da986-1159">function</span></span>||<span data-ttu-id="da986-1160">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da986-1160">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="da986-1161">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="da986-1161">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="da986-1162">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="da986-1162">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="da986-1163">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-1163">Object</span></span>|<span data-ttu-id="da986-1164">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1164">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-1165">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-1165">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="da986-1166">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-1166">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-1167">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-1167">Requirements</span></span>

|<span data-ttu-id="da986-1168">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-1168">Requirement</span></span>|<span data-ttu-id="da986-1169">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-1169">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-1170">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-1170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-1171">1.0</span><span class="sxs-lookup"><span data-stu-id="da986-1171">1.0</span></span>|
|[<span data-ttu-id="da986-1172">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-1172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-1173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-1173">ReadItem</span></span>|
|[<span data-ttu-id="da986-1174">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-1174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-1175">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-1175">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-1176">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-1176">Example</span></span>

<span data-ttu-id="da986-p176">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="da986-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="da986-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="da986-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="da986-1181">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="da986-1181">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="da986-p177">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="da986-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-1186">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-1186">Parameters</span></span>

|<span data-ttu-id="da986-1187">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-1187">Name</span></span>|<span data-ttu-id="da986-1188">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-1188">Type</span></span>|<span data-ttu-id="da986-1189">Atributos</span><span class="sxs-lookup"><span data-stu-id="da986-1189">Attributes</span></span>|<span data-ttu-id="da986-1190">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-1190">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="da986-1191">String</span><span class="sxs-lookup"><span data-stu-id="da986-1191">String</span></span>||<span data-ttu-id="da986-1192">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="da986-1192">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="da986-1193">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-1193">Object</span></span>|<span data-ttu-id="da986-1194">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1194">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-1195">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="da986-1195">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="da986-1196">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-1196">Object</span></span>|<span data-ttu-id="da986-1197">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1197">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-1198">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-1198">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="da986-1199">function</span><span class="sxs-lookup"><span data-stu-id="da986-1199">function</span></span>|<span data-ttu-id="da986-1200">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1200">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-1201">Quando o método for concluído, a função transmitida ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da986-1201">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="da986-1202">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="da986-1202">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="da986-1203">Erros</span><span class="sxs-lookup"><span data-stu-id="da986-1203">Errors</span></span>

|<span data-ttu-id="da986-1204">Código de erro</span><span class="sxs-lookup"><span data-stu-id="da986-1204">Error code</span></span>|<span data-ttu-id="da986-1205">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-1205">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="da986-1206">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="da986-1206">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-1207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-1207">Requirements</span></span>

|<span data-ttu-id="da986-1208">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-1208">Requirement</span></span>|<span data-ttu-id="da986-1209">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-1209">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-1210">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-1210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-1211">1.1</span><span class="sxs-lookup"><span data-stu-id="da986-1211">1.1</span></span>|
|[<span data-ttu-id="da986-1212">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-1212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-1213">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da986-1213">ReadWriteItem</span></span>|
|[<span data-ttu-id="da986-1214">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-1214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-1215">Escrever</span><span class="sxs-lookup"><span data-stu-id="da986-1215">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-1216">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-1216">Example</span></span>

<span data-ttu-id="da986-1217">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="da986-1217">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="da986-1218">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="da986-1218">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="da986-1219">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="da986-1219">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="da986-1220">Atualmente, os tipos de evento compatíveis são `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="da986-1220">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-1221">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-1221">Parameters</span></span>

| <span data-ttu-id="da986-1222">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-1222">Name</span></span> | <span data-ttu-id="da986-1223">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-1223">Type</span></span> | <span data-ttu-id="da986-1224">Atributos</span><span class="sxs-lookup"><span data-stu-id="da986-1224">Attributes</span></span> | <span data-ttu-id="da986-1225">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-1225">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="da986-1226">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="da986-1226">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="da986-1227">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="da986-1227">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="da986-1228">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-1228">Object</span></span> | <span data-ttu-id="da986-1229">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="da986-1230">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="da986-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="da986-1231">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-1231">Object</span></span> | <span data-ttu-id="da986-1232">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="da986-1233">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="da986-1234">function</span><span class="sxs-lookup"><span data-stu-id="da986-1234">function</span></span>| <span data-ttu-id="da986-1235">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-1236">Quando o método for concluído, a função transmitida ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da986-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-1237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-1237">Requirements</span></span>

|<span data-ttu-id="da986-1238">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-1238">Requirement</span></span>| <span data-ttu-id="da986-1239">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-1240">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-1240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da986-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="da986-1241">1.7</span></span> |
|[<span data-ttu-id="da986-1242">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da986-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da986-1243">ReadItem</span></span> |
|[<span data-ttu-id="da986-1244">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="da986-1245">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="da986-1245">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="da986-1246">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-1246">Example</span></span>

```javascript
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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="da986-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="da986-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="da986-1248">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="da986-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="da986-p178">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="da986-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-1252">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="da986-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="da986-1253">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="da986-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="da986-p180">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="da986-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="da986-1257">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="da986-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="da986-1258">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="da986-1258">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="da986-1259">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="da986-1259">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="da986-1260">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="da986-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-1261">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-1261">Parameters</span></span>

|<span data-ttu-id="da986-1262">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-1262">Name</span></span>|<span data-ttu-id="da986-1263">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-1263">Type</span></span>|<span data-ttu-id="da986-1264">Atributos</span><span class="sxs-lookup"><span data-stu-id="da986-1264">Attributes</span></span>|<span data-ttu-id="da986-1265">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="da986-1266">Object</span><span class="sxs-lookup"><span data-stu-id="da986-1266">Object</span></span>|<span data-ttu-id="da986-1267">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-1268">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="da986-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="da986-1269">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-1269">Object</span></span>|<span data-ttu-id="da986-1270">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-1271">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="da986-1272">function</span><span class="sxs-lookup"><span data-stu-id="da986-1272">function</span></span>||<span data-ttu-id="da986-1273">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da986-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="da986-1274">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="da986-1274">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-1275">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-1275">Requirements</span></span>

|<span data-ttu-id="da986-1276">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-1276">Requirement</span></span>|<span data-ttu-id="da986-1277">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-1278">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-1278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="da986-1279">1.3</span></span>|
|[<span data-ttu-id="da986-1280">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da986-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="da986-1282">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-1283">Escrever</span><span class="sxs-lookup"><span data-stu-id="da986-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="da986-1284">Exemplos</span><span class="sxs-lookup"><span data-stu-id="da986-1284">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="da986-p182">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="da986-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="da986-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="da986-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="da986-1288">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="da986-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="da986-p183">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="da986-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="da986-1292">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="da986-1292">Parameters</span></span>

|<span data-ttu-id="da986-1293">Nome</span><span class="sxs-lookup"><span data-stu-id="da986-1293">Name</span></span>|<span data-ttu-id="da986-1294">Tipo</span><span class="sxs-lookup"><span data-stu-id="da986-1294">Type</span></span>|<span data-ttu-id="da986-1295">Atributos</span><span class="sxs-lookup"><span data-stu-id="da986-1295">Attributes</span></span>|<span data-ttu-id="da986-1296">Descrição</span><span class="sxs-lookup"><span data-stu-id="da986-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="da986-1297">String</span><span class="sxs-lookup"><span data-stu-id="da986-1297">String</span></span>||<span data-ttu-id="da986-p184">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="da986-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="da986-1301">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-1301">Object</span></span>|<span data-ttu-id="da986-1302">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-1303">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="da986-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="da986-1304">Objeto</span><span class="sxs-lookup"><span data-stu-id="da986-1304">Object</span></span>|<span data-ttu-id="da986-1305">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-1306">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="da986-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="da986-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="da986-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="da986-1308">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="da986-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="da986-p185">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="da986-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="da986-p186">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="da986-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="da986-1313">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="da986-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="da986-1314">function</span><span class="sxs-lookup"><span data-stu-id="da986-1314">function</span></span>||<span data-ttu-id="da986-1315">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="da986-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="da986-1316">Requisitos</span><span class="sxs-lookup"><span data-stu-id="da986-1316">Requirements</span></span>

|<span data-ttu-id="da986-1317">Requisito</span><span class="sxs-lookup"><span data-stu-id="da986-1317">Requirement</span></span>|<span data-ttu-id="da986-1318">Valor</span><span class="sxs-lookup"><span data-stu-id="da986-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="da986-1319">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="da986-1319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="da986-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="da986-1320">1.2</span></span>|
|[<span data-ttu-id="da986-1321">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="da986-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="da986-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="da986-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="da986-1323">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="da986-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="da986-1324">Escrever</span><span class="sxs-lookup"><span data-stu-id="da986-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="da986-1325">Exemplo</span><span class="sxs-lookup"><span data-stu-id="da986-1325">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
