---
title: Office.Context.Mailbox.item - requisito definir 1.7
description: ''
ms.date: 01/16/2019
localization_priority: Normal
ms.openlocfilehash: dfc86d8a118ab5f5c32968c567a2eec6b9e7d267
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389582"
---
# <a name="item"></a><span data-ttu-id="893bd-102">item</span><span class="sxs-lookup"><span data-stu-id="893bd-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="893bd-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="893bd-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="893bd-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="893bd-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-106">Requirements</span></span>

|<span data-ttu-id="893bd-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-107">Requirement</span></span>|<span data-ttu-id="893bd-108">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-110">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-110">1.0</span></span>|
|[<span data-ttu-id="893bd-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="893bd-112">Restricted</span></span>|
|[<span data-ttu-id="893bd-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-114">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="893bd-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="893bd-115">Members and methods</span></span>

| <span data-ttu-id="893bd-116">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-116">Member</span></span> | <span data-ttu-id="893bd-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="893bd-118">attachments</span><span class="sxs-lookup"><span data-stu-id="893bd-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="893bd-119">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-119">Member</span></span> |
| [<span data-ttu-id="893bd-120">bcc</span><span class="sxs-lookup"><span data-stu-id="893bd-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="893bd-121">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-121">Member</span></span> |
| [<span data-ttu-id="893bd-122">body</span><span class="sxs-lookup"><span data-stu-id="893bd-122">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="893bd-123">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-123">Member</span></span> |
| [<span data-ttu-id="893bd-124">cc</span><span class="sxs-lookup"><span data-stu-id="893bd-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="893bd-125">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-125">Member</span></span> |
| [<span data-ttu-id="893bd-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="893bd-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="893bd-127">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-127">Member</span></span> |
| [<span data-ttu-id="893bd-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="893bd-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="893bd-129">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-129">Member</span></span> |
| [<span data-ttu-id="893bd-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="893bd-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="893bd-131">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-131">Member</span></span> |
| [<span data-ttu-id="893bd-132">end</span><span class="sxs-lookup"><span data-stu-id="893bd-132">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="893bd-133">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-133">Member</span></span> |
| [<span data-ttu-id="893bd-134">from</span><span class="sxs-lookup"><span data-stu-id="893bd-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="893bd-135">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-135">Member</span></span> |
| [<span data-ttu-id="893bd-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="893bd-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="893bd-137">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-137">Member</span></span> |
| [<span data-ttu-id="893bd-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="893bd-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="893bd-139">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-139">Member</span></span> |
| [<span data-ttu-id="893bd-140">itemId</span><span class="sxs-lookup"><span data-stu-id="893bd-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="893bd-141">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-141">Member</span></span> |
| [<span data-ttu-id="893bd-142">itemType</span><span class="sxs-lookup"><span data-stu-id="893bd-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="893bd-143">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-143">Member</span></span> |
| [<span data-ttu-id="893bd-144">location</span><span class="sxs-lookup"><span data-stu-id="893bd-144">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="893bd-145">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-145">Member</span></span> |
| [<span data-ttu-id="893bd-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="893bd-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="893bd-147">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-147">Member</span></span> |
| [<span data-ttu-id="893bd-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="893bd-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="893bd-149">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-149">Member</span></span> |
| [<span data-ttu-id="893bd-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="893bd-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="893bd-151">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-151">Member</span></span> |
| [<span data-ttu-id="893bd-152">organizer</span><span class="sxs-lookup"><span data-stu-id="893bd-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="893bd-153">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-153">Member</span></span> |
| [<span data-ttu-id="893bd-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="893bd-154">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="893bd-155">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-155">Member</span></span> |
| [<span data-ttu-id="893bd-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="893bd-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="893bd-157">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-157">Member</span></span> |
| [<span data-ttu-id="893bd-158">sender</span><span class="sxs-lookup"><span data-stu-id="893bd-158">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="893bd-159">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-159">Member</span></span> |
| [<span data-ttu-id="893bd-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="893bd-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="893bd-161">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-161">Member</span></span> |
| [<span data-ttu-id="893bd-162">start</span><span class="sxs-lookup"><span data-stu-id="893bd-162">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="893bd-163">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-163">Member</span></span> |
| [<span data-ttu-id="893bd-164">subject</span><span class="sxs-lookup"><span data-stu-id="893bd-164">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="893bd-165">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-165">Member</span></span> |
| [<span data-ttu-id="893bd-166">to</span><span class="sxs-lookup"><span data-stu-id="893bd-166">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="893bd-167">Membro</span><span class="sxs-lookup"><span data-stu-id="893bd-167">Member</span></span> |
| [<span data-ttu-id="893bd-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="893bd-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="893bd-169">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-169">Method</span></span> |
| [<span data-ttu-id="893bd-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="893bd-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="893bd-171">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-171">Method</span></span> |
| [<span data-ttu-id="893bd-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="893bd-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="893bd-173">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-173">Method</span></span> |
| [<span data-ttu-id="893bd-174">close</span><span class="sxs-lookup"><span data-stu-id="893bd-174">close</span></span>](#close) | <span data-ttu-id="893bd-175">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-175">Method</span></span> |
| [<span data-ttu-id="893bd-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="893bd-176">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="893bd-177">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-177">Method</span></span> |
| [<span data-ttu-id="893bd-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="893bd-178">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="893bd-179">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-179">Method</span></span> |
| [<span data-ttu-id="893bd-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="893bd-180">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="893bd-181">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-181">Method</span></span> |
| [<span data-ttu-id="893bd-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="893bd-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="893bd-183">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-183">Method</span></span> |
| [<span data-ttu-id="893bd-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="893bd-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="893bd-185">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-185">Method</span></span> |
| [<span data-ttu-id="893bd-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="893bd-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="893bd-187">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-187">Method</span></span> |
| [<span data-ttu-id="893bd-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="893bd-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="893bd-189">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-189">Method</span></span> |
| [<span data-ttu-id="893bd-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="893bd-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="893bd-191">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-191">Method</span></span> |
| [<span data-ttu-id="893bd-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="893bd-192">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="893bd-193">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-193">Method</span></span> |
| [<span data-ttu-id="893bd-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="893bd-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="893bd-195">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-195">Method</span></span> |
| [<span data-ttu-id="893bd-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="893bd-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="893bd-197">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-197">Method</span></span> |
| [<span data-ttu-id="893bd-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="893bd-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="893bd-199">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-199">Method</span></span> |
| [<span data-ttu-id="893bd-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="893bd-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="893bd-201">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-201">Method</span></span> |
| [<span data-ttu-id="893bd-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="893bd-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="893bd-203">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-203">Method</span></span> |
| [<span data-ttu-id="893bd-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="893bd-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="893bd-205">Método</span><span class="sxs-lookup"><span data-stu-id="893bd-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="893bd-206">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-206">Example</span></span>

<span data-ttu-id="893bd-207">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="893bd-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="893bd-208">Membros</span><span class="sxs-lookup"><span data-stu-id="893bd-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="893bd-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="893bd-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="893bd-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="893bd-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-212">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="893bd-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="893bd-213">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="893bd-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-214">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-214">Type:</span></span>

*   <span data-ttu-id="893bd-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="893bd-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-216">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-216">Requirements</span></span>

|<span data-ttu-id="893bd-217">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-217">Requirement</span></span>|<span data-ttu-id="893bd-218">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-219">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-220">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-220">1.0</span></span>|
|[<span data-ttu-id="893bd-221">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-222">ReadItem</span></span>|
|[<span data-ttu-id="893bd-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-224">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-225">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-225">Example</span></span>

<span data-ttu-id="893bd-226">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="893bd-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="893bd-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="893bd-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="893bd-228">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="893bd-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="893bd-229">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="893bd-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-230">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-230">Type:</span></span>

*   [<span data-ttu-id="893bd-231">Destinatários</span><span class="sxs-lookup"><span data-stu-id="893bd-231">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="893bd-232">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-232">Requirements</span></span>

|<span data-ttu-id="893bd-233">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-233">Requirement</span></span>|<span data-ttu-id="893bd-234">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-235">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-236">1.1</span><span class="sxs-lookup"><span data-stu-id="893bd-236">1.1</span></span>|
|[<span data-ttu-id="893bd-237">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-237">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-238">ReadItem</span></span>|
|[<span data-ttu-id="893bd-239">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-239">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-240">Escrever</span><span class="sxs-lookup"><span data-stu-id="893bd-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-241">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-241">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="893bd-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="893bd-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="893bd-243">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="893bd-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-244">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-244">Type:</span></span>

*   [<span data-ttu-id="893bd-245">Corpo</span><span class="sxs-lookup"><span data-stu-id="893bd-245">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="893bd-246">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-246">Requirements</span></span>

|<span data-ttu-id="893bd-247">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-247">Requirement</span></span>|<span data-ttu-id="893bd-248">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-249">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-250">1.1</span><span class="sxs-lookup"><span data-stu-id="893bd-250">1.1</span></span>|
|[<span data-ttu-id="893bd-251">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-252">ReadItem</span></span>|
|[<span data-ttu-id="893bd-253">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-254">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-254">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="893bd-255">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="893bd-255">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="893bd-256">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="893bd-256">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="893bd-257">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="893bd-257">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="893bd-258">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-258">Read mode</span></span>

<span data-ttu-id="893bd-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="893bd-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="893bd-261">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="893bd-261">Compose mode</span></span>

<span data-ttu-id="893bd-262">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="893bd-262">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-263">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-263">Type:</span></span>

*   <span data-ttu-id="893bd-264">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="893bd-264">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-265">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-265">Requirements</span></span>

|<span data-ttu-id="893bd-266">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-266">Requirement</span></span>|<span data-ttu-id="893bd-267">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-268">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-269">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-269">1.0</span></span>|
|[<span data-ttu-id="893bd-270">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-270">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-271">ReadItem</span></span>|
|[<span data-ttu-id="893bd-272">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-272">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-273">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-273">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-274">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-274">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="893bd-275">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="893bd-275">(nullable) conversationId :String</span></span>

<span data-ttu-id="893bd-276">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="893bd-276">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="893bd-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="893bd-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="893bd-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="893bd-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-281">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-281">Type:</span></span>

*   <span data-ttu-id="893bd-282">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="893bd-282">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-283">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-283">Requirements</span></span>

|<span data-ttu-id="893bd-284">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-284">Requirement</span></span>|<span data-ttu-id="893bd-285">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-285">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-286">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-286">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-287">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-287">1.0</span></span>|
|[<span data-ttu-id="893bd-288">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-288">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-289">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-289">ReadItem</span></span>|
|[<span data-ttu-id="893bd-290">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-290">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-291">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-291">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="893bd-292">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="893bd-292">dateTimeCreated :Date</span></span>

<span data-ttu-id="893bd-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="893bd-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-295">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-295">Type:</span></span>

*   <span data-ttu-id="893bd-296">Data</span><span class="sxs-lookup"><span data-stu-id="893bd-296">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-297">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-297">Requirements</span></span>

|<span data-ttu-id="893bd-298">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-298">Requirement</span></span>|<span data-ttu-id="893bd-299">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-300">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-301">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-301">1.0</span></span>|
|[<span data-ttu-id="893bd-302">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-302">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-303">ReadItem</span></span>|
|[<span data-ttu-id="893bd-304">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-304">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-305">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-305">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-306">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-306">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="893bd-307">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="893bd-307">dateTimeModified :Date</span></span>

<span data-ttu-id="893bd-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="893bd-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-310">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="893bd-310">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-311">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-311">Type:</span></span>

*   <span data-ttu-id="893bd-312">Data</span><span class="sxs-lookup"><span data-stu-id="893bd-312">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-313">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-313">Requirements</span></span>

|<span data-ttu-id="893bd-314">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-314">Requirement</span></span>|<span data-ttu-id="893bd-315">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-316">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-317">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-317">1.0</span></span>|
|[<span data-ttu-id="893bd-318">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-318">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-319">ReadItem</span></span>|
|[<span data-ttu-id="893bd-320">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-320">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-321">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-321">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-322">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-322">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="893bd-323">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="893bd-323">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="893bd-324">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="893bd-324">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="893bd-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="893bd-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="893bd-327">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-327">Read mode</span></span>

<span data-ttu-id="893bd-328">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="893bd-328">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="893bd-329">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="893bd-329">Compose mode</span></span>

<span data-ttu-id="893bd-330">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="893bd-330">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="893bd-331">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="893bd-331">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-332">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-332">Type:</span></span>

*   <span data-ttu-id="893bd-333">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="893bd-333">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-334">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-334">Requirements</span></span>

|<span data-ttu-id="893bd-335">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-335">Requirement</span></span>|<span data-ttu-id="893bd-336">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-337">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-338">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-338">1.0</span></span>|
|[<span data-ttu-id="893bd-339">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-339">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-340">ReadItem</span></span>|
|[<span data-ttu-id="893bd-341">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-341">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-342">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-342">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-343">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-343">Example</span></span>

<span data-ttu-id="893bd-344">O exemplo a seguir define a hora de término de um compromisso no modo de redação usando o método [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="893bd-344">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="893bd-345">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="893bd-345">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="893bd-346">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="893bd-346">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="893bd-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="893bd-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-349">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="893bd-349">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="893bd-350">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-350">Read mode</span></span>

<span data-ttu-id="893bd-351">A propriedade `from` retorna um objeto `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="893bd-351">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="893bd-352">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="893bd-352">Compose mode</span></span>

<span data-ttu-id="893bd-353">A propriedade `from` retorna um objeto `From` que fornece um método para obtenção do valor de from.</span><span class="sxs-lookup"><span data-stu-id="893bd-353">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="893bd-354">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-354">Type:</span></span>

*   <span data-ttu-id="893bd-355">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="893bd-355">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-356">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-356">Requirements</span></span>

|<span data-ttu-id="893bd-357">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-357">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="893bd-358">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-359">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-359">1.0</span></span>|<span data-ttu-id="893bd-360">1.7</span><span class="sxs-lookup"><span data-stu-id="893bd-360">1.7</span></span>|
|[<span data-ttu-id="893bd-361">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-362">ReadItem</span></span>|<span data-ttu-id="893bd-363">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="893bd-363">ReadWriteItem</span></span>|
|[<span data-ttu-id="893bd-364">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-365">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-365">Read</span></span>|<span data-ttu-id="893bd-366">Escrever</span><span class="sxs-lookup"><span data-stu-id="893bd-366">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="893bd-367">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="893bd-367">internetMessageId :String</span></span>

<span data-ttu-id="893bd-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="893bd-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-370">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-370">Type:</span></span>

*   <span data-ttu-id="893bd-371">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="893bd-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-372">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-372">Requirements</span></span>

|<span data-ttu-id="893bd-373">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-373">Requirement</span></span>|<span data-ttu-id="893bd-374">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-375">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-376">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-376">1.0</span></span>|
|[<span data-ttu-id="893bd-377">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-378">ReadItem</span></span>|
|[<span data-ttu-id="893bd-379">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-380">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-381">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-381">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="893bd-382">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="893bd-382">itemClass :String</span></span>

<span data-ttu-id="893bd-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="893bd-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="893bd-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="893bd-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="893bd-387">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-387">Type</span></span>|<span data-ttu-id="893bd-388">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-388">Description</span></span>|<span data-ttu-id="893bd-389">classe de item</span><span class="sxs-lookup"><span data-stu-id="893bd-389">item class</span></span>|
|---|---|---|
|<span data-ttu-id="893bd-390">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="893bd-390">Appointment items</span></span>|<span data-ttu-id="893bd-391">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="893bd-391">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="893bd-392">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="893bd-392">Message items</span></span>|<span data-ttu-id="893bd-393">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="893bd-393">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="893bd-394">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="893bd-394">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-395">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-395">Type:</span></span>

*   <span data-ttu-id="893bd-396">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="893bd-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-397">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-397">Requirements</span></span>

|<span data-ttu-id="893bd-398">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-398">Requirement</span></span>|<span data-ttu-id="893bd-399">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-400">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-401">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-401">1.0</span></span>|
|[<span data-ttu-id="893bd-402">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-402">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-403">ReadItem</span></span>|
|[<span data-ttu-id="893bd-404">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-404">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-405">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-406">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-406">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="893bd-407">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="893bd-407">(nullable) itemId :String</span></span>

<span data-ttu-id="893bd-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="893bd-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-410">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="893bd-410">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="893bd-411">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="893bd-411">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="893bd-412">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="893bd-412">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="893bd-413">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="893bd-413">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="893bd-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="893bd-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-416">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-416">Type:</span></span>

*   <span data-ttu-id="893bd-417">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="893bd-417">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-418">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-418">Requirements</span></span>

|<span data-ttu-id="893bd-419">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-419">Requirement</span></span>|<span data-ttu-id="893bd-420">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-421">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-422">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-422">1.0</span></span>|
|[<span data-ttu-id="893bd-423">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-424">ReadItem</span></span>|
|[<span data-ttu-id="893bd-425">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-426">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-426">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-427">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-427">Example</span></span>

<span data-ttu-id="893bd-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="893bd-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="893bd-430">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="893bd-430">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="893bd-431">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="893bd-431">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="893bd-432">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="893bd-432">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-433">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-433">Type:</span></span>

*   [<span data-ttu-id="893bd-434">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="893bd-434">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="893bd-435">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-435">Requirements</span></span>

|<span data-ttu-id="893bd-436">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-436">Requirement</span></span>|<span data-ttu-id="893bd-437">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-438">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-439">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-439">1.0</span></span>|
|[<span data-ttu-id="893bd-440">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-441">ReadItem</span></span>|
|[<span data-ttu-id="893bd-442">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-443">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-444">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-444">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="893bd-445">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="893bd-445">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="893bd-446">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="893bd-446">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="893bd-447">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-447">Read mode</span></span>

<span data-ttu-id="893bd-448">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="893bd-448">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="893bd-449">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="893bd-449">Compose mode</span></span>

<span data-ttu-id="893bd-450">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="893bd-450">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-451">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-451">Type:</span></span>

*   <span data-ttu-id="893bd-452">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="893bd-452">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-453">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-453">Requirements</span></span>

|<span data-ttu-id="893bd-454">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-454">Requirement</span></span>|<span data-ttu-id="893bd-455">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-456">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-457">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-457">1.0</span></span>|
|[<span data-ttu-id="893bd-458">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-458">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-459">ReadItem</span></span>|
|[<span data-ttu-id="893bd-460">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-460">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-461">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-461">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-462">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-462">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="893bd-463">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="893bd-463">normalizedSubject :String</span></span>

<span data-ttu-id="893bd-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="893bd-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="893bd-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject).</span><span class="sxs-lookup"><span data-stu-id="893bd-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-468">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-468">Type:</span></span>

*   <span data-ttu-id="893bd-469">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="893bd-469">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-470">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-470">Requirements</span></span>

|<span data-ttu-id="893bd-471">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-471">Requirement</span></span>|<span data-ttu-id="893bd-472">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-473">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-473">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-474">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-474">1.0</span></span>|
|[<span data-ttu-id="893bd-475">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-476">ReadItem</span></span>|
|[<span data-ttu-id="893bd-477">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-478">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-478">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-479">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-479">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="893bd-480">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="893bd-480">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="893bd-481">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="893bd-481">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-482">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-482">Type:</span></span>

*   [<span data-ttu-id="893bd-483">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="893bd-483">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="893bd-484">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-484">Requirements</span></span>

|<span data-ttu-id="893bd-485">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-485">Requirement</span></span>|<span data-ttu-id="893bd-486">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-487">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-488">1.3</span><span class="sxs-lookup"><span data-stu-id="893bd-488">1.3</span></span>|
|[<span data-ttu-id="893bd-489">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-490">ReadItem</span></span>|
|[<span data-ttu-id="893bd-491">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-492">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-492">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="893bd-493">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="893bd-493">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="893bd-494">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="893bd-494">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="893bd-495">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="893bd-495">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="893bd-496">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-496">Read mode</span></span>

<span data-ttu-id="893bd-497">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="893bd-497">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="893bd-498">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="893bd-498">Compose mode</span></span>

<span data-ttu-id="893bd-499">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="893bd-499">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-500">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-500">Type:</span></span>

*   <span data-ttu-id="893bd-501">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="893bd-501">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-502">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-502">Requirements</span></span>

|<span data-ttu-id="893bd-503">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-503">Requirement</span></span>|<span data-ttu-id="893bd-504">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-505">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-506">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-506">1.0</span></span>|
|[<span data-ttu-id="893bd-507">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-507">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-508">ReadItem</span></span>|
|[<span data-ttu-id="893bd-509">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-509">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-510">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-510">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-511">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-511">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="893bd-512">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="893bd-512">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="893bd-513">Obtém o endereço de email do organizador para uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="893bd-513">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="893bd-514">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-514">Read mode</span></span>

<span data-ttu-id="893bd-515">A propriedade `organizer` retorna um objeto [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="893bd-515">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="893bd-516">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="893bd-516">Compose mode</span></span>

<span data-ttu-id="893bd-517">A propriedade `organizer` retorna um objeto [Organizer](/javascript/api/outlook_1_7/office.organizer) que fornece um método para obtenção do valor de organizer.</span><span class="sxs-lookup"><span data-stu-id="893bd-517">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-518">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-518">Type:</span></span>

*   <span data-ttu-id="893bd-519">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="893bd-519">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-520">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-520">Requirements</span></span>

|<span data-ttu-id="893bd-521">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-521">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="893bd-522">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-523">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-523">1.0</span></span>|<span data-ttu-id="893bd-524">1.7</span><span class="sxs-lookup"><span data-stu-id="893bd-524">1.7</span></span>|
|[<span data-ttu-id="893bd-525">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-525">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-526">ReadItem</span></span>|<span data-ttu-id="893bd-527">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="893bd-527">ReadWriteItem</span></span>|
|[<span data-ttu-id="893bd-528">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-528">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-529">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-529">Read</span></span>|<span data-ttu-id="893bd-530">Escrever</span><span class="sxs-lookup"><span data-stu-id="893bd-530">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-531">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-531">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="893bd-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="893bd-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="893bd-533">Obtém ou configura o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="893bd-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="893bd-534">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="893bd-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="893bd-535">Modos de leitura e redação para itens do compromisso.</span><span class="sxs-lookup"><span data-stu-id="893bd-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="893bd-536">Modo de leitura para os itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="893bd-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="893bd-537">A propriedade `recurrence` retorna um objeto [recurrence](/javascript/api/outlook_1_7/office.recurrence) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="893bd-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="893bd-538">`null` retorna para compromissos individuais e solicitações de reunião de compromissos individuais.</span><span class="sxs-lookup"><span data-stu-id="893bd-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="893bd-539">`undefined` retorna para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="893bd-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="893bd-540">Observação: solicitações de reunião têm um valor `itemClass` de IPM.Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="893bd-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="893bd-541">Observação: se o objeto de recorrência for `null`, isso indicará que o objeto é um compromisso individual ou uma solicitação de reunião de um compromisso individual e NÃO parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="893bd-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-542">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-542">Type:</span></span>

* [<span data-ttu-id="893bd-543">Recurrence</span><span class="sxs-lookup"><span data-stu-id="893bd-543">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="893bd-544">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-544">Requirement</span></span>|<span data-ttu-id="893bd-545">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-546">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-547">1.7</span><span class="sxs-lookup"><span data-stu-id="893bd-547">1.7</span></span>|
|[<span data-ttu-id="893bd-548">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-548">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-549">ReadItem</span></span>|
|[<span data-ttu-id="893bd-550">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-550">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-551">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-551">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="893bd-552">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="893bd-552">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="893bd-553">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="893bd-553">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="893bd-554">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="893bd-554">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="893bd-555">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-555">Read mode</span></span>

<span data-ttu-id="893bd-556">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="893bd-556">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="893bd-557">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="893bd-557">Compose mode</span></span>

<span data-ttu-id="893bd-558">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="893bd-558">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-559">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-559">Type:</span></span>

*   <span data-ttu-id="893bd-560">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="893bd-560">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-561">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-561">Requirements</span></span>

|<span data-ttu-id="893bd-562">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-562">Requirement</span></span>|<span data-ttu-id="893bd-563">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-564">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-565">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-565">1.0</span></span>|
|[<span data-ttu-id="893bd-566">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-567">ReadItem</span></span>|
|[<span data-ttu-id="893bd-568">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-569">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-570">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-570">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="893bd-571">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="893bd-571">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="893bd-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="893bd-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="893bd-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="893bd-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-576">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="893bd-576">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-577">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-577">Type:</span></span>

*   [<span data-ttu-id="893bd-578">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="893bd-578">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="893bd-579">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-579">Requirements</span></span>

|<span data-ttu-id="893bd-580">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-580">Requirement</span></span>|<span data-ttu-id="893bd-581">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-582">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-583">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-583">1.0</span></span>|
|[<span data-ttu-id="893bd-584">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-585">ReadItem</span></span>|
|[<span data-ttu-id="893bd-586">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-587">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-587">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-588">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-588">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="893bd-589">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="893bd-589">(nullable) seriesId :String</span></span>

<span data-ttu-id="893bd-590">Obtém a id da série a qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="893bd-590">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="893bd-591">No OWA e no Outlook, o `seriesId` retorna a ID dos Serviços Web do Exchange (EWS) do item pai (série) a qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="893bd-591">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="893bd-592">No entanto, no iOS e no Android, o `seriesId` retorna a ID REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="893bd-592">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-593">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="893bd-593">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="893bd-594">A propriedade `seriesId` não é idêntica à ID do Outlook usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="893bd-594">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="893bd-595">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="893bd-595">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="893bd-596">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="893bd-596">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="893bd-597">A propriedade `seriesId` retorna `null` para itens que não têm itens pai como compromissos individuais, itens de série ou solicitações de reunião e retorna `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="893bd-597">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-598">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-598">Type:</span></span>

* <span data-ttu-id="893bd-599">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="893bd-599">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-600">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-600">Requirements</span></span>

|<span data-ttu-id="893bd-601">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-601">Requirement</span></span>|<span data-ttu-id="893bd-602">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-603">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-604">1.7</span><span class="sxs-lookup"><span data-stu-id="893bd-604">1.7</span></span>|
|[<span data-ttu-id="893bd-605">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-605">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-606">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-606">ReadItem</span></span>|
|[<span data-ttu-id="893bd-607">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-607">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-608">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-608">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-609">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-609">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="893bd-610">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="893bd-610">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="893bd-611">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="893bd-611">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="893bd-p130">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="893bd-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="893bd-614">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-614">Read mode</span></span>

<span data-ttu-id="893bd-615">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="893bd-615">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="893bd-616">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="893bd-616">Compose mode</span></span>

<span data-ttu-id="893bd-617">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="893bd-617">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="893bd-618">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="893bd-618">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-619">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-619">Type:</span></span>

*   <span data-ttu-id="893bd-620">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="893bd-620">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-621">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-621">Requirements</span></span>

|<span data-ttu-id="893bd-622">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-622">Requirement</span></span>|<span data-ttu-id="893bd-623">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-624">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-625">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-625">1.0</span></span>|
|[<span data-ttu-id="893bd-626">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-627">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-627">ReadItem</span></span>|
|[<span data-ttu-id="893bd-628">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-629">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-629">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-630">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-630">Example</span></span>

<span data-ttu-id="893bd-631">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="893bd-631">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="893bd-632">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="893bd-632">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="893bd-633">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="893bd-633">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="893bd-634">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="893bd-634">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="893bd-635">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-635">Read mode</span></span>

<span data-ttu-id="893bd-p131">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="893bd-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="893bd-638">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="893bd-638">Compose mode</span></span>

<span data-ttu-id="893bd-639">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="893bd-639">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="893bd-640">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-640">Type:</span></span>

*   <span data-ttu-id="893bd-641">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="893bd-641">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-642">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-642">Requirements</span></span>

|<span data-ttu-id="893bd-643">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-643">Requirement</span></span>|<span data-ttu-id="893bd-644">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-644">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-645">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-645">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-646">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-646">1.0</span></span>|
|[<span data-ttu-id="893bd-647">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-647">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-648">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-648">ReadItem</span></span>|
|[<span data-ttu-id="893bd-649">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-649">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-650">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-650">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="893bd-651">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="893bd-651">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="893bd-652">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="893bd-652">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="893bd-653">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="893bd-653">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="893bd-654">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-654">Read mode</span></span>

<span data-ttu-id="893bd-p133">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="893bd-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="893bd-657">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="893bd-657">Compose mode</span></span>

<span data-ttu-id="893bd-658">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="893bd-658">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="893bd-659">Tipo:</span><span class="sxs-lookup"><span data-stu-id="893bd-659">Type:</span></span>

*   <span data-ttu-id="893bd-660">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="893bd-660">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-661">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-661">Requirements</span></span>

|<span data-ttu-id="893bd-662">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-662">Requirement</span></span>|<span data-ttu-id="893bd-663">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-663">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-664">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-664">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-665">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-665">1.0</span></span>|
|[<span data-ttu-id="893bd-666">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-666">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-667">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-667">ReadItem</span></span>|
|[<span data-ttu-id="893bd-668">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-668">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-669">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-669">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-670">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-670">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="893bd-671">Métodos</span><span class="sxs-lookup"><span data-stu-id="893bd-671">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="893bd-672">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="893bd-672">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="893bd-673">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="893bd-673">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="893bd-674">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="893bd-674">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="893bd-675">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="893bd-675">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-676">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-676">Parameters:</span></span>
|<span data-ttu-id="893bd-677">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-677">Name</span></span>|<span data-ttu-id="893bd-678">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-678">Type</span></span>|<span data-ttu-id="893bd-679">Atributos</span><span class="sxs-lookup"><span data-stu-id="893bd-679">Attributes</span></span>|<span data-ttu-id="893bd-680">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-680">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="893bd-681">String</span><span class="sxs-lookup"><span data-stu-id="893bd-681">String</span></span>||<span data-ttu-id="893bd-p134">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="893bd-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="893bd-684">String</span><span class="sxs-lookup"><span data-stu-id="893bd-684">String</span></span>||<span data-ttu-id="893bd-p135">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="893bd-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="893bd-687">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-687">Object</span></span>|<span data-ttu-id="893bd-688">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-688">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-689">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="893bd-689">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="893bd-690">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-690">Object</span></span>|<span data-ttu-id="893bd-691">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-691">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-692">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="893bd-692">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="893bd-693">Booliano</span><span class="sxs-lookup"><span data-stu-id="893bd-693">Boolean</span></span>|<span data-ttu-id="893bd-694">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-694">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-695">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="893bd-695">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="893bd-696">function</span><span class="sxs-lookup"><span data-stu-id="893bd-696">function</span></span>|<span data-ttu-id="893bd-697">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-697">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-698">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="893bd-698">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="893bd-699">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="893bd-699">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="893bd-700">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="893bd-700">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="893bd-701">Erros</span><span class="sxs-lookup"><span data-stu-id="893bd-701">Errors</span></span>

|<span data-ttu-id="893bd-702">Código de erro</span><span class="sxs-lookup"><span data-stu-id="893bd-702">Error code</span></span>|<span data-ttu-id="893bd-703">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-703">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="893bd-704">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="893bd-704">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="893bd-705">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="893bd-705">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="893bd-706">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="893bd-706">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-707">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-707">Requirements</span></span>

|<span data-ttu-id="893bd-708">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-708">Requirement</span></span>|<span data-ttu-id="893bd-709">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-709">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-710">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-710">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-711">1.1</span><span class="sxs-lookup"><span data-stu-id="893bd-711">1.1</span></span>|
|[<span data-ttu-id="893bd-712">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-712">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-713">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="893bd-713">ReadWriteItem</span></span>|
|[<span data-ttu-id="893bd-714">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-714">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-715">Escrever</span><span class="sxs-lookup"><span data-stu-id="893bd-715">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="893bd-716">Exemplos</span><span class="sxs-lookup"><span data-stu-id="893bd-716">Examples</span></span>

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

<span data-ttu-id="893bd-717">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="893bd-717">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="893bd-718">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="893bd-718">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="893bd-719">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="893bd-719">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="893bd-720">Atualmente, os tipos de evento compatíveis são `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="893bd-720">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-721">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-721">Parameters:</span></span>

| <span data-ttu-id="893bd-722">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-722">Name</span></span> | <span data-ttu-id="893bd-723">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-723">Type</span></span> | <span data-ttu-id="893bd-724">Atributos</span><span class="sxs-lookup"><span data-stu-id="893bd-724">Attributes</span></span> | <span data-ttu-id="893bd-725">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-725">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="893bd-726">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="893bd-726">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="893bd-727">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="893bd-727">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="893bd-728">Função</span><span class="sxs-lookup"><span data-stu-id="893bd-728">Function</span></span> || <span data-ttu-id="893bd-p136">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="893bd-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="893bd-732">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-732">Object</span></span> | <span data-ttu-id="893bd-733">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-733">&lt;optional&gt;</span></span> | <span data-ttu-id="893bd-734">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="893bd-734">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="893bd-735">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-735">Object</span></span> | <span data-ttu-id="893bd-736">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-736">&lt;optional&gt;</span></span> | <span data-ttu-id="893bd-737">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="893bd-737">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="893bd-738">function</span><span class="sxs-lookup"><span data-stu-id="893bd-738">function</span></span>| <span data-ttu-id="893bd-739">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-739">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-740">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="893bd-740">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-741">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-741">Requirements</span></span>

|<span data-ttu-id="893bd-742">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-742">Requirement</span></span>| <span data-ttu-id="893bd-743">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-743">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-744">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-744">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="893bd-745">1.7</span><span class="sxs-lookup"><span data-stu-id="893bd-745">1.7</span></span> |
|[<span data-ttu-id="893bd-746">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-746">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="893bd-747">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-747">ReadItem</span></span> |
|[<span data-ttu-id="893bd-748">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-748">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="893bd-749">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-749">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="893bd-750">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-750">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="893bd-751">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="893bd-751">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="893bd-752">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="893bd-752">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="893bd-p137">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="893bd-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="893bd-756">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="893bd-756">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="893bd-757">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="893bd-757">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-758">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-758">Parameters:</span></span>

|<span data-ttu-id="893bd-759">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-759">Name</span></span>|<span data-ttu-id="893bd-760">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-760">Type</span></span>|<span data-ttu-id="893bd-761">Atributos</span><span class="sxs-lookup"><span data-stu-id="893bd-761">Attributes</span></span>|<span data-ttu-id="893bd-762">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-762">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="893bd-763">String</span><span class="sxs-lookup"><span data-stu-id="893bd-763">String</span></span>||<span data-ttu-id="893bd-p138">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="893bd-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="893bd-766">String</span><span class="sxs-lookup"><span data-stu-id="893bd-766">String</span></span>||<span data-ttu-id="893bd-p139">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="893bd-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="893bd-769">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-769">Object</span></span>|<span data-ttu-id="893bd-770">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-770">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-771">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="893bd-771">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="893bd-772">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-772">Object</span></span>|<span data-ttu-id="893bd-773">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-773">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-774">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="893bd-774">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="893bd-775">function</span><span class="sxs-lookup"><span data-stu-id="893bd-775">function</span></span>|<span data-ttu-id="893bd-776">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-776">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-777">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="893bd-777">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="893bd-778">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="893bd-778">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="893bd-779">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="893bd-779">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="893bd-780">Erros</span><span class="sxs-lookup"><span data-stu-id="893bd-780">Errors</span></span>

|<span data-ttu-id="893bd-781">Código de erro</span><span class="sxs-lookup"><span data-stu-id="893bd-781">Error code</span></span>|<span data-ttu-id="893bd-782">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-782">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="893bd-783">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="893bd-783">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-784">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-784">Requirements</span></span>

|<span data-ttu-id="893bd-785">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-785">Requirement</span></span>|<span data-ttu-id="893bd-786">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-786">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-787">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-787">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-788">1.1</span><span class="sxs-lookup"><span data-stu-id="893bd-788">1.1</span></span>|
|[<span data-ttu-id="893bd-789">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-789">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-790">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="893bd-790">ReadWriteItem</span></span>|
|[<span data-ttu-id="893bd-791">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-791">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-792">Escrever</span><span class="sxs-lookup"><span data-stu-id="893bd-792">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-793">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-793">Example</span></span>

<span data-ttu-id="893bd-794">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="893bd-794">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
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

####  <a name="close"></a><span data-ttu-id="893bd-795">close()</span><span class="sxs-lookup"><span data-stu-id="893bd-795">close()</span></span>

<span data-ttu-id="893bd-796">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="893bd-796">Closes the current item that is being composed.</span></span>

<span data-ttu-id="893bd-p140">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="893bd-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-799">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="893bd-799">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="893bd-800">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="893bd-800">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-801">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-801">Requirements</span></span>

|<span data-ttu-id="893bd-802">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-802">Requirement</span></span>|<span data-ttu-id="893bd-803">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-803">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-804">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-804">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-805">1.3</span><span class="sxs-lookup"><span data-stu-id="893bd-805">1.3</span></span>|
|[<span data-ttu-id="893bd-806">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-806">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-807">Restrito</span><span class="sxs-lookup"><span data-stu-id="893bd-807">Restricted</span></span>|
|[<span data-ttu-id="893bd-808">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-808">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-809">Escrever</span><span class="sxs-lookup"><span data-stu-id="893bd-809">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="893bd-810">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="893bd-810">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="893bd-811">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="893bd-811">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-812">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="893bd-812">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="893bd-813">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="893bd-813">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="893bd-814">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="893bd-814">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="893bd-p141">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="893bd-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-818">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-818">Parameters:</span></span>

|<span data-ttu-id="893bd-819">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-819">Name</span></span>|<span data-ttu-id="893bd-820">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-820">Type</span></span>|<span data-ttu-id="893bd-821">Atributos</span><span class="sxs-lookup"><span data-stu-id="893bd-821">Attributes</span></span>|<span data-ttu-id="893bd-822">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-822">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="893bd-823">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="893bd-823">String &#124; Object</span></span>||<span data-ttu-id="893bd-p142">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="893bd-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="893bd-826">**OU**</span><span class="sxs-lookup"><span data-stu-id="893bd-826">**OR**</span></span><br/><span data-ttu-id="893bd-p143">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="893bd-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="893bd-829">String</span><span class="sxs-lookup"><span data-stu-id="893bd-829">String</span></span>|<span data-ttu-id="893bd-830">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-830">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="893bd-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="893bd-833">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-833">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="893bd-834">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-834">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-835">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="893bd-835">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="893bd-836">String</span><span class="sxs-lookup"><span data-stu-id="893bd-836">String</span></span>||<span data-ttu-id="893bd-p145">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="893bd-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="893bd-839">String</span><span class="sxs-lookup"><span data-stu-id="893bd-839">String</span></span>||<span data-ttu-id="893bd-840">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="893bd-840">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="893bd-841">String</span><span class="sxs-lookup"><span data-stu-id="893bd-841">String</span></span>||<span data-ttu-id="893bd-p146">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="893bd-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="893bd-844">Booliano</span><span class="sxs-lookup"><span data-stu-id="893bd-844">Boolean</span></span>||<span data-ttu-id="893bd-p147">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="893bd-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="893bd-847">String</span><span class="sxs-lookup"><span data-stu-id="893bd-847">String</span></span>||<span data-ttu-id="893bd-p148">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="893bd-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="893bd-851">function</span><span class="sxs-lookup"><span data-stu-id="893bd-851">function</span></span>|<span data-ttu-id="893bd-852">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-852">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-853">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="893bd-853">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-854">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-854">Requirements</span></span>

|<span data-ttu-id="893bd-855">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-855">Requirement</span></span>|<span data-ttu-id="893bd-856">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-856">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-857">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-857">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-858">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-858">1.0</span></span>|
|[<span data-ttu-id="893bd-859">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-859">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-860">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-860">ReadItem</span></span>|
|[<span data-ttu-id="893bd-861">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-861">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-862">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-862">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="893bd-863">Exemplos</span><span class="sxs-lookup"><span data-stu-id="893bd-863">Examples</span></span>

<span data-ttu-id="893bd-864">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="893bd-864">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="893bd-865">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="893bd-865">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="893bd-866">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="893bd-866">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="893bd-867">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="893bd-867">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="893bd-868">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="893bd-868">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="893bd-869">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="893bd-869">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="893bd-870">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="893bd-870">displayReplyForm(formData)</span></span>

<span data-ttu-id="893bd-871">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="893bd-871">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-872">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="893bd-872">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="893bd-873">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="893bd-873">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="893bd-874">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="893bd-874">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="893bd-p149">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="893bd-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-878">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-878">Parameters:</span></span>

|<span data-ttu-id="893bd-879">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-879">Name</span></span>|<span data-ttu-id="893bd-880">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-880">Type</span></span>|<span data-ttu-id="893bd-881">Atributos</span><span class="sxs-lookup"><span data-stu-id="893bd-881">Attributes</span></span>|<span data-ttu-id="893bd-882">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-882">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="893bd-883">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="893bd-883">String &#124; Object</span></span>||<span data-ttu-id="893bd-p150">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="893bd-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="893bd-886">**OU**</span><span class="sxs-lookup"><span data-stu-id="893bd-886">**OR**</span></span><br/><span data-ttu-id="893bd-p151">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="893bd-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="893bd-889">String</span><span class="sxs-lookup"><span data-stu-id="893bd-889">String</span></span>|<span data-ttu-id="893bd-890">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-890">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-p152">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="893bd-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="893bd-893">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-893">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="893bd-894">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-894">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-895">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="893bd-895">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="893bd-896">String</span><span class="sxs-lookup"><span data-stu-id="893bd-896">String</span></span>||<span data-ttu-id="893bd-p153">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="893bd-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="893bd-899">String</span><span class="sxs-lookup"><span data-stu-id="893bd-899">String</span></span>||<span data-ttu-id="893bd-900">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="893bd-900">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="893bd-901">String</span><span class="sxs-lookup"><span data-stu-id="893bd-901">String</span></span>||<span data-ttu-id="893bd-p154">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="893bd-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="893bd-904">Booliano</span><span class="sxs-lookup"><span data-stu-id="893bd-904">Boolean</span></span>||<span data-ttu-id="893bd-p155">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="893bd-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="893bd-907">String</span><span class="sxs-lookup"><span data-stu-id="893bd-907">String</span></span>||<span data-ttu-id="893bd-p156">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="893bd-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="893bd-911">function</span><span class="sxs-lookup"><span data-stu-id="893bd-911">function</span></span>|<span data-ttu-id="893bd-912">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-912">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-913">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="893bd-913">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-914">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-914">Requirements</span></span>

|<span data-ttu-id="893bd-915">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-915">Requirement</span></span>|<span data-ttu-id="893bd-916">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-916">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-917">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-917">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-918">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-918">1.0</span></span>|
|[<span data-ttu-id="893bd-919">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-919">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-920">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-920">ReadItem</span></span>|
|[<span data-ttu-id="893bd-921">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-921">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-922">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-922">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="893bd-923">Exemplos</span><span class="sxs-lookup"><span data-stu-id="893bd-923">Examples</span></span>

<span data-ttu-id="893bd-924">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="893bd-924">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="893bd-925">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="893bd-925">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="893bd-926">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="893bd-926">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="893bd-927">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="893bd-927">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="893bd-928">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="893bd-928">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="893bd-929">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="893bd-929">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="893bd-930">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="893bd-930">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="893bd-931">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="893bd-931">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-932">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="893bd-932">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-933">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-933">Requirements</span></span>

|<span data-ttu-id="893bd-934">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-934">Requirement</span></span>|<span data-ttu-id="893bd-935">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-935">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-936">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-936">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-937">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-937">1.0</span></span>|
|[<span data-ttu-id="893bd-938">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-938">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-939">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-939">ReadItem</span></span>|
|[<span data-ttu-id="893bd-940">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-940">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-941">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-941">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="893bd-942">Retorna:</span><span class="sxs-lookup"><span data-stu-id="893bd-942">Returns:</span></span>

<span data-ttu-id="893bd-943">Tipo: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="893bd-943">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="893bd-944">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-944">Example</span></span>

<span data-ttu-id="893bd-945">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="893bd-945">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="893bd-946">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="893bd-946">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="893bd-947">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="893bd-947">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-948">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="893bd-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-949">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-949">Parameters:</span></span>

|<span data-ttu-id="893bd-950">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-950">Name</span></span>|<span data-ttu-id="893bd-951">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-951">Type</span></span>|<span data-ttu-id="893bd-952">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-952">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="893bd-953">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="893bd-953">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="893bd-954">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="893bd-954">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-955">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-955">Requirements</span></span>

|<span data-ttu-id="893bd-956">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-956">Requirement</span></span>|<span data-ttu-id="893bd-957">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-957">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-958">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-958">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-959">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-959">1.0</span></span>|
|[<span data-ttu-id="893bd-960">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-960">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-961">Restrito</span><span class="sxs-lookup"><span data-stu-id="893bd-961">Restricted</span></span>|
|[<span data-ttu-id="893bd-962">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-962">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-963">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-963">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="893bd-964">Retorna:</span><span class="sxs-lookup"><span data-stu-id="893bd-964">Returns:</span></span>

<span data-ttu-id="893bd-965">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="893bd-965">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="893bd-966">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="893bd-966">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="893bd-967">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="893bd-967">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="893bd-968">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="893bd-968">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="893bd-969">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="893bd-969">Value of `entityType`</span></span>|<span data-ttu-id="893bd-970">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="893bd-970">Type of objects in returned array</span></span>|<span data-ttu-id="893bd-971">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="893bd-971">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="893bd-972">String</span><span class="sxs-lookup"><span data-stu-id="893bd-972">String</span></span>|<span data-ttu-id="893bd-973">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="893bd-973">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="893bd-974">Contato</span><span class="sxs-lookup"><span data-stu-id="893bd-974">Contact</span></span>|<span data-ttu-id="893bd-975">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="893bd-975">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="893bd-976">String</span><span class="sxs-lookup"><span data-stu-id="893bd-976">String</span></span>|<span data-ttu-id="893bd-977">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="893bd-977">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="893bd-978">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="893bd-978">MeetingSuggestion</span></span>|<span data-ttu-id="893bd-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="893bd-979">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="893bd-980">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="893bd-980">PhoneNumber</span></span>|<span data-ttu-id="893bd-981">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="893bd-981">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="893bd-982">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="893bd-982">TaskSuggestion</span></span>|<span data-ttu-id="893bd-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="893bd-983">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="893bd-984">String</span><span class="sxs-lookup"><span data-stu-id="893bd-984">String</span></span>|<span data-ttu-id="893bd-985">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="893bd-985">**Restricted**</span></span>|

<span data-ttu-id="893bd-986">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="893bd-986">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="893bd-987">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-987">Example</span></span>

<span data-ttu-id="893bd-988">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="893bd-988">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="893bd-989">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="893bd-989">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="893bd-990">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="893bd-990">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-991">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="893bd-991">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="893bd-992">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="893bd-992">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-993">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-993">Parameters:</span></span>

|<span data-ttu-id="893bd-994">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-994">Name</span></span>|<span data-ttu-id="893bd-995">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-995">Type</span></span>|<span data-ttu-id="893bd-996">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-996">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="893bd-997">String</span><span class="sxs-lookup"><span data-stu-id="893bd-997">String</span></span>|<span data-ttu-id="893bd-998">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="893bd-998">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-999">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-999">Requirements</span></span>

|<span data-ttu-id="893bd-1000">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-1000">Requirement</span></span>|<span data-ttu-id="893bd-1001">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-1002">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-1003">1.0</span></span>|
|[<span data-ttu-id="893bd-1004">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-1004">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-1005">ReadItem</span></span>|
|[<span data-ttu-id="893bd-1006">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-1006">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-1007">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-1007">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="893bd-1008">Retorna:</span><span class="sxs-lookup"><span data-stu-id="893bd-1008">Returns:</span></span>

<span data-ttu-id="893bd-p158">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="893bd-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="893bd-1011">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="893bd-1011">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="893bd-1012">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="893bd-1012">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="893bd-1013">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="893bd-1013">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-1014">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="893bd-1014">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="893bd-p159">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="893bd-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="893bd-1018">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="893bd-1018">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="893bd-1019">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="893bd-1019">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="893bd-p160">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="893bd-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-1023">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-1023">Requirements</span></span>

|<span data-ttu-id="893bd-1024">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-1024">Requirement</span></span>|<span data-ttu-id="893bd-1025">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-1026">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-1027">1.0</span></span>|
|[<span data-ttu-id="893bd-1028">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-1029">ReadItem</span></span>|
|[<span data-ttu-id="893bd-1030">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-1031">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-1031">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="893bd-1032">Retorna:</span><span class="sxs-lookup"><span data-stu-id="893bd-1032">Returns:</span></span>

<span data-ttu-id="893bd-p161">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="893bd-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="893bd-1035">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="893bd-1035">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="893bd-1036">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-1036">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="893bd-1037">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-1037">Example</span></span>

<span data-ttu-id="893bd-1038">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="893bd-1038">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="893bd-1039">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="893bd-1039">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="893bd-1040">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="893bd-1040">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-1041">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="893bd-1041">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="893bd-1042">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="893bd-1042">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="893bd-p162">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="893bd-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-1045">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-1045">Parameters:</span></span>

|<span data-ttu-id="893bd-1046">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-1046">Name</span></span>|<span data-ttu-id="893bd-1047">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-1047">Type</span></span>|<span data-ttu-id="893bd-1048">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-1048">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="893bd-1049">String</span><span class="sxs-lookup"><span data-stu-id="893bd-1049">String</span></span>|<span data-ttu-id="893bd-1050">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="893bd-1050">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-1051">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-1051">Requirements</span></span>

|<span data-ttu-id="893bd-1052">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-1052">Requirement</span></span>|<span data-ttu-id="893bd-1053">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-1053">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-1054">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-1054">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-1055">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-1055">1.0</span></span>|
|[<span data-ttu-id="893bd-1056">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-1056">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-1057">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-1057">ReadItem</span></span>|
|[<span data-ttu-id="893bd-1058">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-1058">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-1059">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-1059">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="893bd-1060">Retorna:</span><span class="sxs-lookup"><span data-stu-id="893bd-1060">Returns:</span></span>

<span data-ttu-id="893bd-1061">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="893bd-1061">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="893bd-1062">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="893bd-1062">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="893bd-1063">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="893bd-1063">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="893bd-1064">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-1064">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="893bd-1065">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="893bd-1065">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="893bd-1066">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="893bd-1066">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="893bd-p163">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="893bd-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-1069">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-1069">Parameters:</span></span>

|<span data-ttu-id="893bd-1070">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-1070">Name</span></span>|<span data-ttu-id="893bd-1071">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-1071">Type</span></span>|<span data-ttu-id="893bd-1072">Atributos</span><span class="sxs-lookup"><span data-stu-id="893bd-1072">Attributes</span></span>|<span data-ttu-id="893bd-1073">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-1073">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="893bd-1074">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="893bd-1074">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="893bd-p164">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="893bd-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="893bd-1078">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-1078">Object</span></span>|<span data-ttu-id="893bd-1079">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1079">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-1080">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="893bd-1080">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="893bd-1081">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-1081">Object</span></span>|<span data-ttu-id="893bd-1082">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1082">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-1083">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="893bd-1083">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="893bd-1084">function</span><span class="sxs-lookup"><span data-stu-id="893bd-1084">function</span></span>||<span data-ttu-id="893bd-1085">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="893bd-1085">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="893bd-1086">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="893bd-1086">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="893bd-1087">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="893bd-1087">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-1088">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-1088">Requirements</span></span>

|<span data-ttu-id="893bd-1089">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-1089">Requirement</span></span>|<span data-ttu-id="893bd-1090">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-1091">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-1092">1.2</span><span class="sxs-lookup"><span data-stu-id="893bd-1092">1.2</span></span>|
|[<span data-ttu-id="893bd-1093">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-1094">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="893bd-1094">ReadWriteItem</span></span>|
|[<span data-ttu-id="893bd-1095">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-1096">Escrever</span><span class="sxs-lookup"><span data-stu-id="893bd-1096">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="893bd-1097">Retorna:</span><span class="sxs-lookup"><span data-stu-id="893bd-1097">Returns:</span></span>

<span data-ttu-id="893bd-1098">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="893bd-1098">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="893bd-1099">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="893bd-1099">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="893bd-1100">String</span><span class="sxs-lookup"><span data-stu-id="893bd-1100">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="893bd-1101">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-1101">Example</span></span>

```js
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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="893bd-1102">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="893bd-1102">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="893bd-p166">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="893bd-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-1105">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="893bd-1105">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-1106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-1106">Requirements</span></span>

|<span data-ttu-id="893bd-1107">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-1107">Requirement</span></span>|<span data-ttu-id="893bd-1108">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-1108">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-1109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-1109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-1110">1.6</span><span class="sxs-lookup"><span data-stu-id="893bd-1110">1.6</span></span>|
|[<span data-ttu-id="893bd-1111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-1111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-1112">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-1112">ReadItem</span></span>|
|[<span data-ttu-id="893bd-1113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-1113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-1114">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-1114">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="893bd-1115">Retorna:</span><span class="sxs-lookup"><span data-stu-id="893bd-1115">Returns:</span></span>

<span data-ttu-id="893bd-1116">Tipo: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="893bd-1116">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="893bd-1117">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-1117">Example</span></span>

<span data-ttu-id="893bd-1118">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="893bd-1118">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="893bd-1119">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="893bd-1119">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="893bd-p167">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="893bd-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-1122">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="893bd-1122">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="893bd-p168">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="893bd-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="893bd-1126">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="893bd-1126">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="893bd-1127">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="893bd-1127">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="893bd-p169">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="893bd-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="893bd-1131">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-1131">Requirements</span></span>

|<span data-ttu-id="893bd-1132">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-1132">Requirement</span></span>|<span data-ttu-id="893bd-1133">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-1133">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-1134">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-1134">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-1135">1.6</span><span class="sxs-lookup"><span data-stu-id="893bd-1135">1.6</span></span>|
|[<span data-ttu-id="893bd-1136">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-1136">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-1137">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-1137">ReadItem</span></span>|
|[<span data-ttu-id="893bd-1138">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-1138">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-1139">Read</span><span class="sxs-lookup"><span data-stu-id="893bd-1139">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="893bd-1140">Retorna:</span><span class="sxs-lookup"><span data-stu-id="893bd-1140">Returns:</span></span>

<span data-ttu-id="893bd-p170">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="893bd-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="893bd-1143">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-1143">Example</span></span>

<span data-ttu-id="893bd-1144">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="893bd-1144">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="893bd-1145">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="893bd-1145">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="893bd-1146">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="893bd-1146">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="893bd-p171">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="893bd-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-1150">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-1150">Parameters:</span></span>

|<span data-ttu-id="893bd-1151">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-1151">Name</span></span>|<span data-ttu-id="893bd-1152">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-1152">Type</span></span>|<span data-ttu-id="893bd-1153">Atributos</span><span class="sxs-lookup"><span data-stu-id="893bd-1153">Attributes</span></span>|<span data-ttu-id="893bd-1154">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-1154">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="893bd-1155">function</span><span class="sxs-lookup"><span data-stu-id="893bd-1155">function</span></span>||<span data-ttu-id="893bd-1156">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="893bd-1156">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="893bd-1157">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="893bd-1157">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="893bd-1158">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="893bd-1158">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="893bd-1159">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-1159">Object</span></span>|<span data-ttu-id="893bd-1160">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-1161">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="893bd-1161">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="893bd-1162">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="893bd-1162">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-1163">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-1163">Requirements</span></span>

|<span data-ttu-id="893bd-1164">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-1164">Requirement</span></span>|<span data-ttu-id="893bd-1165">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-1165">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-1166">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-1166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-1167">1.0</span><span class="sxs-lookup"><span data-stu-id="893bd-1167">1.0</span></span>|
|[<span data-ttu-id="893bd-1168">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-1168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-1169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-1169">ReadItem</span></span>|
|[<span data-ttu-id="893bd-1170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-1170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-1171">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-1171">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-1172">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-1172">Example</span></span>

<span data-ttu-id="893bd-p174">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="893bd-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="893bd-1176">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="893bd-1176">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="893bd-1177">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="893bd-1177">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="893bd-p175">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="893bd-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-1182">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-1182">Parameters:</span></span>

|<span data-ttu-id="893bd-1183">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-1183">Name</span></span>|<span data-ttu-id="893bd-1184">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-1184">Type</span></span>|<span data-ttu-id="893bd-1185">Atributos</span><span class="sxs-lookup"><span data-stu-id="893bd-1185">Attributes</span></span>|<span data-ttu-id="893bd-1186">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-1186">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="893bd-1187">String</span><span class="sxs-lookup"><span data-stu-id="893bd-1187">String</span></span>||<span data-ttu-id="893bd-1188">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="893bd-1188">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="893bd-1189">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-1189">Object</span></span>|<span data-ttu-id="893bd-1190">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-1191">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="893bd-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="893bd-1192">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-1192">Object</span></span>|<span data-ttu-id="893bd-1193">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-1194">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="893bd-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="893bd-1195">function</span><span class="sxs-lookup"><span data-stu-id="893bd-1195">function</span></span>|<span data-ttu-id="893bd-1196">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-1197">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="893bd-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="893bd-1198">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="893bd-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="893bd-1199">Erros</span><span class="sxs-lookup"><span data-stu-id="893bd-1199">Errors</span></span>

|<span data-ttu-id="893bd-1200">Código de erro</span><span class="sxs-lookup"><span data-stu-id="893bd-1200">Error code</span></span>|<span data-ttu-id="893bd-1201">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="893bd-1202">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="893bd-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-1203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-1203">Requirements</span></span>

|<span data-ttu-id="893bd-1204">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-1204">Requirement</span></span>|<span data-ttu-id="893bd-1205">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-1206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-1206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="893bd-1207">1.1</span></span>|
|[<span data-ttu-id="893bd-1208">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="893bd-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="893bd-1210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-1211">Escrever</span><span class="sxs-lookup"><span data-stu-id="893bd-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-1212">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-1212">Example</span></span>

<span data-ttu-id="893bd-1213">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="893bd-1213">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="893bd-1214">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="893bd-1214">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="893bd-1215">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="893bd-1215">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="893bd-1216">Atualmente, os tipos de evento compatíveis são `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="893bd-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-1217">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-1217">Parameters:</span></span>

| <span data-ttu-id="893bd-1218">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-1218">Name</span></span> | <span data-ttu-id="893bd-1219">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-1219">Type</span></span> | <span data-ttu-id="893bd-1220">Atributos</span><span class="sxs-lookup"><span data-stu-id="893bd-1220">Attributes</span></span> | <span data-ttu-id="893bd-1221">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="893bd-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="893bd-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="893bd-1223">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="893bd-1223">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="893bd-1224">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-1224">Object</span></span> | <span data-ttu-id="893bd-1225">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1225">&lt;optional&gt;</span></span> | <span data-ttu-id="893bd-1226">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="893bd-1226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="893bd-1227">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-1227">Object</span></span> | <span data-ttu-id="893bd-1228">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1228">&lt;optional&gt;</span></span> | <span data-ttu-id="893bd-1229">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="893bd-1229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="893bd-1230">function</span><span class="sxs-lookup"><span data-stu-id="893bd-1230">function</span></span>| <span data-ttu-id="893bd-1231">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-1232">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="893bd-1232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-1233">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-1233">Requirements</span></span>

|<span data-ttu-id="893bd-1234">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-1234">Requirement</span></span>| <span data-ttu-id="893bd-1235">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-1235">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-1236">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-1236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="893bd-1237">1.7</span><span class="sxs-lookup"><span data-stu-id="893bd-1237">1.7</span></span> |
|[<span data-ttu-id="893bd-1238">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-1238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="893bd-1239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="893bd-1239">ReadItem</span></span> |
|[<span data-ttu-id="893bd-1240">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-1240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="893bd-1241">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="893bd-1241">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="893bd-1242">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-1242">Example</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="893bd-1243">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="893bd-1243">saveAsync([options], callback)</span></span>

<span data-ttu-id="893bd-1244">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="893bd-1244">Asynchronously saves an item.</span></span>

<span data-ttu-id="893bd-p176">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="893bd-p176">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-1248">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="893bd-1248">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="893bd-1249">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="893bd-1249">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="893bd-p178">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="893bd-p178">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="893bd-1253">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="893bd-1253">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="893bd-1254">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="893bd-1254">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="893bd-1255">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="893bd-1255">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="893bd-1256">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="893bd-1256">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-1257">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-1257">Parameters:</span></span>

|<span data-ttu-id="893bd-1258">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-1258">Name</span></span>|<span data-ttu-id="893bd-1259">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-1259">Type</span></span>|<span data-ttu-id="893bd-1260">Atributos</span><span class="sxs-lookup"><span data-stu-id="893bd-1260">Attributes</span></span>|<span data-ttu-id="893bd-1261">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-1261">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="893bd-1262">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-1262">Object</span></span>|<span data-ttu-id="893bd-1263">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1263">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-1264">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="893bd-1264">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="893bd-1265">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-1265">Object</span></span>|<span data-ttu-id="893bd-1266">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1266">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-1267">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="893bd-1267">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="893bd-1268">function</span><span class="sxs-lookup"><span data-stu-id="893bd-1268">function</span></span>||<span data-ttu-id="893bd-1269">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="893bd-1269">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="893bd-1270">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="893bd-1270">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-1271">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-1271">Requirements</span></span>

|<span data-ttu-id="893bd-1272">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-1272">Requirement</span></span>|<span data-ttu-id="893bd-1273">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-1273">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-1274">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-1274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-1275">1.3</span><span class="sxs-lookup"><span data-stu-id="893bd-1275">1.3</span></span>|
|[<span data-ttu-id="893bd-1276">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-1276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-1277">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="893bd-1277">ReadWriteItem</span></span>|
|[<span data-ttu-id="893bd-1278">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-1278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-1279">Escrever</span><span class="sxs-lookup"><span data-stu-id="893bd-1279">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="893bd-1280">Exemplos</span><span class="sxs-lookup"><span data-stu-id="893bd-1280">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="893bd-p180">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="893bd-p180">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="893bd-1283">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="893bd-1283">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="893bd-1284">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="893bd-1284">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="893bd-p181">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="893bd-p181">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="893bd-1288">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="893bd-1288">Parameters:</span></span>

|<span data-ttu-id="893bd-1289">Nome</span><span class="sxs-lookup"><span data-stu-id="893bd-1289">Name</span></span>|<span data-ttu-id="893bd-1290">Tipo</span><span class="sxs-lookup"><span data-stu-id="893bd-1290">Type</span></span>|<span data-ttu-id="893bd-1291">Atributos</span><span class="sxs-lookup"><span data-stu-id="893bd-1291">Attributes</span></span>|<span data-ttu-id="893bd-1292">Descrição</span><span class="sxs-lookup"><span data-stu-id="893bd-1292">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="893bd-1293">String</span><span class="sxs-lookup"><span data-stu-id="893bd-1293">String</span></span>||<span data-ttu-id="893bd-p182">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="893bd-p182">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="893bd-1297">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-1297">Object</span></span>|<span data-ttu-id="893bd-1298">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1298">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-1299">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="893bd-1299">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="893bd-1300">Objeto</span><span class="sxs-lookup"><span data-stu-id="893bd-1300">Object</span></span>|<span data-ttu-id="893bd-1301">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1301">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-1302">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="893bd-1302">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="893bd-1303">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="893bd-1303">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="893bd-1304">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="893bd-1304">&lt;optional&gt;</span></span>|<span data-ttu-id="893bd-p183">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="893bd-p183">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="893bd-p184">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="893bd-p184">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="893bd-1309">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="893bd-1309">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="893bd-1310">function</span><span class="sxs-lookup"><span data-stu-id="893bd-1310">function</span></span>||<span data-ttu-id="893bd-1311">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="893bd-1311">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="893bd-1312">Requisitos</span><span class="sxs-lookup"><span data-stu-id="893bd-1312">Requirements</span></span>

|<span data-ttu-id="893bd-1313">Requisito</span><span class="sxs-lookup"><span data-stu-id="893bd-1313">Requirement</span></span>|<span data-ttu-id="893bd-1314">Valor</span><span class="sxs-lookup"><span data-stu-id="893bd-1314">Value</span></span>|
|---|---|
|[<span data-ttu-id="893bd-1315">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="893bd-1315">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="893bd-1316">1.2</span><span class="sxs-lookup"><span data-stu-id="893bd-1316">1.2</span></span>|
|[<span data-ttu-id="893bd-1317">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="893bd-1317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="893bd-1318">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="893bd-1318">ReadWriteItem</span></span>|
|[<span data-ttu-id="893bd-1319">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="893bd-1319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="893bd-1320">Escrever</span><span class="sxs-lookup"><span data-stu-id="893bd-1320">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="893bd-1321">Exemplo</span><span class="sxs-lookup"><span data-stu-id="893bd-1321">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
