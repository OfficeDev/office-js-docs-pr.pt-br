---
title: Office.Context.Mailbox.item - requisito definir 1.6
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: 0c3eca68285e9d415954e6ce45d2a80508fa701b
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701894"
---
# <a name="item"></a><span data-ttu-id="b1ec5-102">item</span><span class="sxs-lookup"><span data-stu-id="b1ec5-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="b1ec5-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="b1ec5-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="b1ec5-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-106">Requirements</span></span>

|<span data-ttu-id="b1ec5-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-107">Requirement</span></span>| <span data-ttu-id="b1ec5-108">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-110">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-110">1.0</span></span>|
|[<span data-ttu-id="b1ec5-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-112">Restricted</span></span>|
|[<span data-ttu-id="b1ec5-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-114">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b1ec5-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-115">Members and methods</span></span>

| <span data-ttu-id="b1ec5-116">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-116">Member</span></span> | <span data-ttu-id="b1ec5-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b1ec5-118">attachments</span><span class="sxs-lookup"><span data-stu-id="b1ec5-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="b1ec5-119">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-119">Member</span></span> |
| [<span data-ttu-id="b1ec5-120">bcc</span><span class="sxs-lookup"><span data-stu-id="b1ec5-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="b1ec5-121">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-121">Member</span></span> |
| [<span data-ttu-id="b1ec5-122">body</span><span class="sxs-lookup"><span data-stu-id="b1ec5-122">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="b1ec5-123">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-123">Member</span></span> |
| [<span data-ttu-id="b1ec5-124">cc</span><span class="sxs-lookup"><span data-stu-id="b1ec5-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="b1ec5-125">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-125">Member</span></span> |
| [<span data-ttu-id="b1ec5-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="b1ec5-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="b1ec5-127">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-127">Member</span></span> |
| [<span data-ttu-id="b1ec5-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="b1ec5-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="b1ec5-129">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-129">Member</span></span> |
| [<span data-ttu-id="b1ec5-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="b1ec5-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="b1ec5-131">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-131">Member</span></span> |
| [<span data-ttu-id="b1ec5-132">end</span><span class="sxs-lookup"><span data-stu-id="b1ec5-132">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="b1ec5-133">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-133">Member</span></span> |
| [<span data-ttu-id="b1ec5-134">from</span><span class="sxs-lookup"><span data-stu-id="b1ec5-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="b1ec5-135">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-135">Member</span></span> |
| [<span data-ttu-id="b1ec5-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="b1ec5-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="b1ec5-137">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-137">Member</span></span> |
| [<span data-ttu-id="b1ec5-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="b1ec5-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="b1ec5-139">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-139">Member</span></span> |
| [<span data-ttu-id="b1ec5-140">itemId</span><span class="sxs-lookup"><span data-stu-id="b1ec5-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="b1ec5-141">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-141">Member</span></span> |
| [<span data-ttu-id="b1ec5-142">itemType</span><span class="sxs-lookup"><span data-stu-id="b1ec5-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="b1ec5-143">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-143">Member</span></span> |
| [<span data-ttu-id="b1ec5-144">location</span><span class="sxs-lookup"><span data-stu-id="b1ec5-144">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="b1ec5-145">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-145">Member</span></span> |
| [<span data-ttu-id="b1ec5-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="b1ec5-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="b1ec5-147">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-147">Member</span></span> |
| [<span data-ttu-id="b1ec5-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="b1ec5-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="b1ec5-149">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-149">Member</span></span> |
| [<span data-ttu-id="b1ec5-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="b1ec5-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="b1ec5-151">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-151">Member</span></span> |
| [<span data-ttu-id="b1ec5-152">organizer</span><span class="sxs-lookup"><span data-stu-id="b1ec5-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="b1ec5-153">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-153">Member</span></span> |
| [<span data-ttu-id="b1ec5-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="b1ec5-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="b1ec5-155">Member</span><span class="sxs-lookup"><span data-stu-id="b1ec5-155">Member</span></span> |
| [<span data-ttu-id="b1ec5-156">sender</span><span class="sxs-lookup"><span data-stu-id="b1ec5-156">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="b1ec5-157">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-157">Member</span></span> |
| [<span data-ttu-id="b1ec5-158">start</span><span class="sxs-lookup"><span data-stu-id="b1ec5-158">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="b1ec5-159">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-159">Member</span></span> |
| [<span data-ttu-id="b1ec5-160">subject</span><span class="sxs-lookup"><span data-stu-id="b1ec5-160">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="b1ec5-161">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-161">Member</span></span> |
| [<span data-ttu-id="b1ec5-162">to</span><span class="sxs-lookup"><span data-stu-id="b1ec5-162">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="b1ec5-163">Membro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-163">Member</span></span> |
| [<span data-ttu-id="b1ec5-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b1ec5-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="b1ec5-165">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-165">Method</span></span> |
| [<span data-ttu-id="b1ec5-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b1ec5-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="b1ec5-167">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-167">Method</span></span> |
| [<span data-ttu-id="b1ec5-168">close</span><span class="sxs-lookup"><span data-stu-id="b1ec5-168">close</span></span>](#close) | <span data-ttu-id="b1ec5-169">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-169">Method</span></span> |
| [<span data-ttu-id="b1ec5-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="b1ec5-170">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="b1ec5-171">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-171">Method</span></span> |
| [<span data-ttu-id="b1ec5-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="b1ec5-172">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="b1ec5-173">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-173">Method</span></span> |
| [<span data-ttu-id="b1ec5-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="b1ec5-174">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="b1ec5-175">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-175">Method</span></span> |
| [<span data-ttu-id="b1ec5-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="b1ec5-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="b1ec5-177">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-177">Method</span></span> |
| [<span data-ttu-id="b1ec5-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="b1ec5-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="b1ec5-179">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-179">Method</span></span> |
| [<span data-ttu-id="b1ec5-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="b1ec5-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="b1ec5-181">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-181">Method</span></span> |
| [<span data-ttu-id="b1ec5-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="b1ec5-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="b1ec5-183">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-183">Method</span></span> |
| [<span data-ttu-id="b1ec5-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="b1ec5-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="b1ec5-185">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-185">Method</span></span> |
| [<span data-ttu-id="b1ec5-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="b1ec5-186">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="b1ec5-187">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-187">Method</span></span> |
| [<span data-ttu-id="b1ec5-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="b1ec5-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="b1ec5-189">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-189">Method</span></span> |
| [<span data-ttu-id="b1ec5-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="b1ec5-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="b1ec5-191">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-191">Method</span></span> |
| [<span data-ttu-id="b1ec5-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b1ec5-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="b1ec5-193">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-193">Method</span></span> |
| [<span data-ttu-id="b1ec5-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="b1ec5-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="b1ec5-195">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-195">Method</span></span> |
| [<span data-ttu-id="b1ec5-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="b1ec5-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="b1ec5-197">Método</span><span class="sxs-lookup"><span data-stu-id="b1ec5-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="b1ec5-198">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-198">Example</span></span>

<span data-ttu-id="b1ec5-199">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="b1ec5-200">Membros</span><span class="sxs-lookup"><span data-stu-id="b1ec5-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="b1ec5-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b1ec5-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="b1ec5-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-204">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="b1ec5-205">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-206">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-206">Type:</span></span>

*   <span data-ttu-id="b1ec5-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b1ec5-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-208">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-208">Requirements</span></span>

|<span data-ttu-id="b1ec5-209">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-209">Requirement</span></span>| <span data-ttu-id="b1ec5-210">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-211">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-212">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-212">1.0</span></span>|
|[<span data-ttu-id="b1ec5-213">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-213">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-214">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-215">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-216">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-217">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-217">Example</span></span>

<span data-ttu-id="b1ec5-218">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="b1ec5-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="b1ec5-220">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="b1ec5-221">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-222">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-222">Type:</span></span>

*   [<span data-ttu-id="b1ec5-223">Destinatários</span><span class="sxs-lookup"><span data-stu-id="b1ec5-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="b1ec5-224">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-224">Requirements</span></span>

|<span data-ttu-id="b1ec5-225">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-225">Requirement</span></span>| <span data-ttu-id="b1ec5-226">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-227">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-228">1.1</span><span class="sxs-lookup"><span data-stu-id="b1ec5-228">1.1</span></span>|
|[<span data-ttu-id="b1ec5-229">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-229">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-230">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-231">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-231">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-232">Escrever</span><span class="sxs-lookup"><span data-stu-id="b1ec5-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-233">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-233">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="b1ec5-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="b1ec5-235">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-236">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-236">Type:</span></span>

*   [<span data-ttu-id="b1ec5-237">Corpo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="b1ec5-238">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-238">Requirements</span></span>

|<span data-ttu-id="b1ec5-239">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-239">Requirement</span></span>| <span data-ttu-id="b1ec5-240">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-241">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-242">1.1</span><span class="sxs-lookup"><span data-stu-id="b1ec5-242">1.1</span></span>|
|[<span data-ttu-id="b1ec5-243">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-243">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-244">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-245">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-245">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-246">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-246">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="b1ec5-247">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-247">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="b1ec5-248">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-248">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="b1ec5-249">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-249">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b1ec5-250">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-250">Read mode</span></span>

<span data-ttu-id="b1ec5-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b1ec5-253">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="b1ec5-253">Compose mode</span></span>

<span data-ttu-id="b1ec5-254">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-254">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-255">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-255">Type:</span></span>

*   <span data-ttu-id="b1ec5-256">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-256">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-257">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-257">Requirements</span></span>

|<span data-ttu-id="b1ec5-258">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-258">Requirement</span></span>| <span data-ttu-id="b1ec5-259">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-260">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-261">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-261">1.0</span></span>|
|[<span data-ttu-id="b1ec5-262">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-262">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-263">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-263">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-264">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-264">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-265">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-265">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-266">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-266">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="b1ec5-267">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-267">(nullable) conversationId :String</span></span>

<span data-ttu-id="b1ec5-268">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-268">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="b1ec5-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="b1ec5-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-273">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-273">Type:</span></span>

*   <span data-ttu-id="b1ec5-274">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b1ec5-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-275">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-275">Requirements</span></span>

|<span data-ttu-id="b1ec5-276">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-276">Requirement</span></span>| <span data-ttu-id="b1ec5-277">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-278">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-279">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-279">1.0</span></span>|
|[<span data-ttu-id="b1ec5-280">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-281">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-282">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-283">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-283">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="b1ec5-284">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="b1ec5-284">dateTimeCreated :Date</span></span>

<span data-ttu-id="b1ec5-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-287">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-287">Type:</span></span>

*   <span data-ttu-id="b1ec5-288">Data</span><span class="sxs-lookup"><span data-stu-id="b1ec5-288">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-289">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-289">Requirements</span></span>

|<span data-ttu-id="b1ec5-290">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-290">Requirement</span></span>| <span data-ttu-id="b1ec5-291">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-291">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-292">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-292">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-293">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-293">1.0</span></span>|
|[<span data-ttu-id="b1ec5-294">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-294">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-295">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-295">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-296">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-296">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-297">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-297">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-298">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-298">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="b1ec5-299">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="b1ec5-299">dateTimeModified :Date</span></span>

<span data-ttu-id="b1ec5-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-302">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-302">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-303">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-303">Type:</span></span>

*   <span data-ttu-id="b1ec5-304">Data</span><span class="sxs-lookup"><span data-stu-id="b1ec5-304">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-305">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-305">Requirements</span></span>

|<span data-ttu-id="b1ec5-306">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-306">Requirement</span></span>| <span data-ttu-id="b1ec5-307">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-308">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-309">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-309">1.0</span></span>|
|[<span data-ttu-id="b1ec5-310">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-310">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-311">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-312">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-312">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-313">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-314">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-314">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="b1ec5-315">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-315">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="b1ec5-316">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-316">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="b1ec5-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b1ec5-319">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-319">Read mode</span></span>

<span data-ttu-id="b1ec5-320">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-320">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b1ec5-321">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="b1ec5-321">Compose mode</span></span>

<span data-ttu-id="b1ec5-322">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-322">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="b1ec5-323">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-323">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-324">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-324">Type:</span></span>

*   <span data-ttu-id="b1ec5-325">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-325">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-326">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-326">Requirements</span></span>

|<span data-ttu-id="b1ec5-327">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-327">Requirement</span></span>| <span data-ttu-id="b1ec5-328">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-329">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-330">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-330">1.0</span></span>|
|[<span data-ttu-id="b1ec5-331">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-331">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-332">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-333">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-333">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-334">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-334">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-335">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-335">Example</span></span>

<span data-ttu-id="b1ec5-336">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-336">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="b1ec5-337">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-337">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="b1ec5-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="b1ec5-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-342">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-342">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-343">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-343">Type:</span></span>

*   [<span data-ttu-id="b1ec5-344">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b1ec5-344">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b1ec5-345">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-345">Requirements</span></span>

|<span data-ttu-id="b1ec5-346">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-346">Requirement</span></span>| <span data-ttu-id="b1ec5-347">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-348">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-349">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-349">1.0</span></span>|
|[<span data-ttu-id="b1ec5-350">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-351">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-352">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-353">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-353">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="b1ec5-354">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-354">internetMessageId :String</span></span>

<span data-ttu-id="b1ec5-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-357">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-357">Type:</span></span>

*   <span data-ttu-id="b1ec5-358">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b1ec5-358">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-359">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-359">Requirements</span></span>

|<span data-ttu-id="b1ec5-360">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-360">Requirement</span></span>| <span data-ttu-id="b1ec5-361">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-362">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-363">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-363">1.0</span></span>|
|[<span data-ttu-id="b1ec5-364">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-365">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-366">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-367">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-367">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-368">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-368">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="b1ec5-369">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-369">itemClass :String</span></span>

<span data-ttu-id="b1ec5-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="b1ec5-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="b1ec5-374">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-374">Type</span></span> | <span data-ttu-id="b1ec5-375">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-375">Description</span></span> | <span data-ttu-id="b1ec5-376">classe de item</span><span class="sxs-lookup"><span data-stu-id="b1ec5-376">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="b1ec5-377">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="b1ec5-377">Appointment items</span></span> | <span data-ttu-id="b1ec5-378">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-378">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="b1ec5-379">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-379">Message items</span></span> | <span data-ttu-id="b1ec5-380">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-380">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="b1ec5-381">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-381">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-382">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-382">Type:</span></span>

*   <span data-ttu-id="b1ec5-383">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b1ec5-383">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-384">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-384">Requirements</span></span>

|<span data-ttu-id="b1ec5-385">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-385">Requirement</span></span>| <span data-ttu-id="b1ec5-386">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-386">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-387">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-387">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-388">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-388">1.0</span></span>|
|[<span data-ttu-id="b1ec5-389">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-389">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-390">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-390">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-391">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-391">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-392">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-392">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-393">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-393">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="b1ec5-394">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-394">(nullable) itemId :String</span></span>

<span data-ttu-id="b1ec5-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-397">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-397">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="b1ec5-398">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-398">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="b1ec5-399">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-399">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="b1ec5-400">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-400">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="b1ec5-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-403">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-403">Type:</span></span>

*   <span data-ttu-id="b1ec5-404">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b1ec5-404">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-405">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-405">Requirements</span></span>

|<span data-ttu-id="b1ec5-406">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-406">Requirement</span></span>| <span data-ttu-id="b1ec5-407">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-408">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-409">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-409">1.0</span></span>|
|[<span data-ttu-id="b1ec5-410">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-411">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-412">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-413">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-413">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-414">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-414">Example</span></span>

<span data-ttu-id="b1ec5-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="b1ec5-417">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-417">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="b1ec5-418">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-418">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="b1ec5-419">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-419">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-420">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-420">Type:</span></span>

*   [<span data-ttu-id="b1ec5-421">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="b1ec5-421">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="b1ec5-422">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-422">Requirements</span></span>

|<span data-ttu-id="b1ec5-423">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-423">Requirement</span></span>| <span data-ttu-id="b1ec5-424">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-424">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-425">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-426">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-426">1.0</span></span>|
|[<span data-ttu-id="b1ec5-427">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-428">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-429">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-430">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-430">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-431">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-431">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="b1ec5-432">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-432">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="b1ec5-433">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-433">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b1ec5-434">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-434">Read mode</span></span>

<span data-ttu-id="b1ec5-435">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-435">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b1ec5-436">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="b1ec5-436">Compose mode</span></span>

<span data-ttu-id="b1ec5-437">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-437">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-438">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-438">Type:</span></span>

*   <span data-ttu-id="b1ec5-439">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-439">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-440">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-440">Requirements</span></span>

|<span data-ttu-id="b1ec5-441">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-441">Requirement</span></span>| <span data-ttu-id="b1ec5-442">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-443">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-443">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-444">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-444">1.0</span></span>|
|[<span data-ttu-id="b1ec5-445">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-446">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-447">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-448">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-449">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-449">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="b1ec5-450">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-450">normalizedSubject :String</span></span>

<span data-ttu-id="b1ec5-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="b1ec5-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-455">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-455">Type:</span></span>

*   <span data-ttu-id="b1ec5-456">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b1ec5-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-457">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-457">Requirements</span></span>

|<span data-ttu-id="b1ec5-458">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-458">Requirement</span></span>| <span data-ttu-id="b1ec5-459">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-460">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-461">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-461">1.0</span></span>|
|[<span data-ttu-id="b1ec5-462">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-463">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-464">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-465">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-466">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="b1ec5-467">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-467">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="b1ec5-468">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-468">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-469">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-469">Type:</span></span>

*   [<span data-ttu-id="b1ec5-470">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="b1ec5-470">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="b1ec5-471">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-471">Requirements</span></span>

|<span data-ttu-id="b1ec5-472">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-472">Requirement</span></span>| <span data-ttu-id="b1ec5-473">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-474">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-475">1.3</span><span class="sxs-lookup"><span data-stu-id="b1ec5-475">1.3</span></span>|
|[<span data-ttu-id="b1ec5-476">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-476">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-477">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-478">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-478">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-479">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-479">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="b1ec5-480">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-480">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="b1ec5-481">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-481">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="b1ec5-482">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-482">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b1ec5-483">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-483">Read mode</span></span>

<span data-ttu-id="b1ec5-484">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-484">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b1ec5-485">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="b1ec5-485">Compose mode</span></span>

<span data-ttu-id="b1ec5-486">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-486">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-487">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-487">Type:</span></span>

*   <span data-ttu-id="b1ec5-488">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-488">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-489">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-489">Requirements</span></span>

|<span data-ttu-id="b1ec5-490">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-490">Requirement</span></span>| <span data-ttu-id="b1ec5-491">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-492">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-492">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-493">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-493">1.0</span></span>|
|[<span data-ttu-id="b1ec5-494">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-495">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-496">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-497">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-497">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-498">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-498">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="b1ec5-499">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-499">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="b1ec5-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-502">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-502">Type:</span></span>

*   [<span data-ttu-id="b1ec5-503">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b1ec5-503">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b1ec5-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-504">Requirements</span></span>

|<span data-ttu-id="b1ec5-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-505">Requirement</span></span>| <span data-ttu-id="b1ec5-506">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-507">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-508">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-508">1.0</span></span>|
|[<span data-ttu-id="b1ec5-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-510">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-511">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-512">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-512">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-513">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-513">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="b1ec5-514">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-514">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="b1ec5-515">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-515">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="b1ec5-516">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-516">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b1ec5-517">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-517">Read mode</span></span>

<span data-ttu-id="b1ec5-518">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-518">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b1ec5-519">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="b1ec5-519">Compose mode</span></span>

<span data-ttu-id="b1ec5-520">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-520">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-521">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-521">Type:</span></span>

*   <span data-ttu-id="b1ec5-522">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-522">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-523">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-523">Requirements</span></span>

|<span data-ttu-id="b1ec5-524">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-524">Requirement</span></span>| <span data-ttu-id="b1ec5-525">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-525">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-526">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-526">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-527">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-527">1.0</span></span>|
|[<span data-ttu-id="b1ec5-528">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-528">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-529">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-530">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-530">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-531">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-531">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-532">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-532">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="b1ec5-533">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-533">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="b1ec5-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="b1ec5-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-538">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-538">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-539">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-539">Type:</span></span>

*   [<span data-ttu-id="b1ec5-540">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b1ec5-540">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b1ec5-541">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-541">Requirements</span></span>

|<span data-ttu-id="b1ec5-542">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-542">Requirement</span></span>| <span data-ttu-id="b1ec5-543">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-543">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-544">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-544">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-545">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-545">1.0</span></span>|
|[<span data-ttu-id="b1ec5-546">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-546">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-547">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-547">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-548">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-548">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-549">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-549">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-550">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-550">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="b1ec5-551">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-551">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="b1ec5-552">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-552">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="b1ec5-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b1ec5-555">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-555">Read mode</span></span>

<span data-ttu-id="b1ec5-556">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-556">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b1ec5-557">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="b1ec5-557">Compose mode</span></span>

<span data-ttu-id="b1ec5-558">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-558">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="b1ec5-559">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-559">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-560">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-560">Type:</span></span>

*   <span data-ttu-id="b1ec5-561">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-561">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-562">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-562">Requirements</span></span>

|<span data-ttu-id="b1ec5-563">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-563">Requirement</span></span>| <span data-ttu-id="b1ec5-564">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-565">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-566">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-566">1.0</span></span>|
|[<span data-ttu-id="b1ec5-567">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-567">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-568">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-568">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-569">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-569">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-570">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-570">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-571">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-571">Example</span></span>

<span data-ttu-id="b1ec5-572">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-572">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="b1ec5-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="b1ec5-574">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="b1ec5-575">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b1ec5-576">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-576">Read mode</span></span>

<span data-ttu-id="b1ec5-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="b1ec5-579">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="b1ec5-579">Compose mode</span></span>

<span data-ttu-id="b1ec5-580">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b1ec5-581">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-581">Type:</span></span>

*   <span data-ttu-id="b1ec5-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-583">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-583">Requirements</span></span>

|<span data-ttu-id="b1ec5-584">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-584">Requirement</span></span>| <span data-ttu-id="b1ec5-585">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-586">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-587">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-587">1.0</span></span>|
|[<span data-ttu-id="b1ec5-588">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-588">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-589">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-590">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-590">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-591">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-591">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="b1ec5-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="b1ec5-593">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="b1ec5-594">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b1ec5-595">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-595">Read mode</span></span>

<span data-ttu-id="b1ec5-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b1ec5-598">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="b1ec5-598">Compose mode</span></span>

<span data-ttu-id="b1ec5-599">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ec5-600">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-600">Type:</span></span>

*   <span data-ttu-id="b1ec5-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-602">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-602">Requirements</span></span>

|<span data-ttu-id="b1ec5-603">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-603">Requirement</span></span>| <span data-ttu-id="b1ec5-604">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-605">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-606">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-606">1.0</span></span>|
|[<span data-ttu-id="b1ec5-607">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-607">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-608">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-609">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-609">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-610">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-610">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-611">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-611">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="b1ec5-612">Métodos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-612">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="b1ec5-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b1ec5-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b1ec5-614">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="b1ec5-615">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="b1ec5-616">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ec5-617">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-617">Parameters:</span></span>

|<span data-ttu-id="b1ec5-618">Nome</span><span class="sxs-lookup"><span data-stu-id="b1ec5-618">Name</span></span>| <span data-ttu-id="b1ec5-619">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-619">Type</span></span>| <span data-ttu-id="b1ec5-620">Atributos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-620">Attributes</span></span>| <span data-ttu-id="b1ec5-621">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="b1ec5-622">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-622">String</span></span>||<span data-ttu-id="b1ec5-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b1ec5-625">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-625">String</span></span>||<span data-ttu-id="b1ec5-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b1ec5-628">Object</span><span class="sxs-lookup"><span data-stu-id="b1ec5-628">Object</span></span>| <span data-ttu-id="b1ec5-629">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-629">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-630">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-630">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="b1ec5-631">Objeto</span><span class="sxs-lookup"><span data-stu-id="b1ec5-631">Object</span></span> | <span data-ttu-id="b1ec5-632">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-632">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ec5-633">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-633">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="b1ec5-634">Booliano</span><span class="sxs-lookup"><span data-stu-id="b1ec5-634">Boolean</span></span> | <span data-ttu-id="b1ec5-635">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-635">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ec5-636">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-636">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="b1ec5-637">function</span><span class="sxs-lookup"><span data-stu-id="b1ec5-637">function</span></span>| <span data-ttu-id="b1ec5-638">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-638">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-639">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b1ec5-640">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-640">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b1ec5-641">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-641">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b1ec5-642">Erros</span><span class="sxs-lookup"><span data-stu-id="b1ec5-642">Errors</span></span>

| <span data-ttu-id="b1ec5-643">Código de erro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-643">Error code</span></span> | <span data-ttu-id="b1ec5-644">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-644">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="b1ec5-645">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-645">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="b1ec5-646">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-646">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b1ec5-647">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-647">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b1ec5-648">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-648">Requirements</span></span>

|<span data-ttu-id="b1ec5-649">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-649">Requirement</span></span>| <span data-ttu-id="b1ec5-650">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-651">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-652">1.1</span><span class="sxs-lookup"><span data-stu-id="b1ec5-652">1.1</span></span>|
|[<span data-ttu-id="b1ec5-653">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-653">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-654">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-654">ReadWriteItem</span></span>|
|[<span data-ttu-id="b1ec5-655">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-655">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-656">Redação</span><span class="sxs-lookup"><span data-stu-id="b1ec5-656">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b1ec5-657">Exemplos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-657">Examples</span></span>

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

<span data-ttu-id="b1ec5-658">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-658">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="b1ec5-659">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b1ec5-659">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b1ec5-660">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-660">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="b1ec5-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="b1ec5-664">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-664">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="b1ec5-665">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-665">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ec5-666">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-666">Parameters:</span></span>

|<span data-ttu-id="b1ec5-667">Nome</span><span class="sxs-lookup"><span data-stu-id="b1ec5-667">Name</span></span>| <span data-ttu-id="b1ec5-668">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-668">Type</span></span>| <span data-ttu-id="b1ec5-669">Atributos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-669">Attributes</span></span>| <span data-ttu-id="b1ec5-670">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-670">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="b1ec5-671">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b1ec5-671">String</span></span>||<span data-ttu-id="b1ec5-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b1ec5-674">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-674">String</span></span>||<span data-ttu-id="b1ec5-p136">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b1ec5-677">Object</span><span class="sxs-lookup"><span data-stu-id="b1ec5-677">Object</span></span>| <span data-ttu-id="b1ec5-678">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-678">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-679">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-679">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b1ec5-680">Objeto</span><span class="sxs-lookup"><span data-stu-id="b1ec5-680">Object</span></span>| <span data-ttu-id="b1ec5-681">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-681">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-682">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-682">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b1ec5-683">function</span><span class="sxs-lookup"><span data-stu-id="b1ec5-683">function</span></span>| <span data-ttu-id="b1ec5-684">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-684">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-685">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-685">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b1ec5-686">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-686">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b1ec5-687">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-687">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b1ec5-688">Erros</span><span class="sxs-lookup"><span data-stu-id="b1ec5-688">Errors</span></span>

| <span data-ttu-id="b1ec5-689">Código de erro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-689">Error code</span></span> | <span data-ttu-id="b1ec5-690">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-690">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b1ec5-691">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-691">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b1ec5-692">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-692">Requirements</span></span>

|<span data-ttu-id="b1ec5-693">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-693">Requirement</span></span>| <span data-ttu-id="b1ec5-694">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-695">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-696">1.1</span><span class="sxs-lookup"><span data-stu-id="b1ec5-696">1.1</span></span>|
|[<span data-ttu-id="b1ec5-697">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-697">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-698">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-698">ReadWriteItem</span></span>|
|[<span data-ttu-id="b1ec5-699">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-699">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-700">Escrever</span><span class="sxs-lookup"><span data-stu-id="b1ec5-700">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-701">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-701">Example</span></span>

<span data-ttu-id="b1ec5-702">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-702">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="b1ec5-703">close()</span><span class="sxs-lookup"><span data-stu-id="b1ec5-703">close()</span></span>

<span data-ttu-id="b1ec5-704">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-704">Closes the current item that is being composed.</span></span>

<span data-ttu-id="b1ec5-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-707">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-707">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="b1ec5-708">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-708">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-709">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-709">Requirements</span></span>

|<span data-ttu-id="b1ec5-710">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-710">Requirement</span></span>| <span data-ttu-id="b1ec5-711">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-711">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-712">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-712">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-713">1.3</span><span class="sxs-lookup"><span data-stu-id="b1ec5-713">1.3</span></span>|
|[<span data-ttu-id="b1ec5-714">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-714">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-715">Restrito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-715">Restricted</span></span>|
|[<span data-ttu-id="b1ec5-716">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-716">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-717">Escrever</span><span class="sxs-lookup"><span data-stu-id="b1ec5-717">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="b1ec5-718">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-718">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="b1ec5-719">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-719">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-720">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-720">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b1ec5-721">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-721">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b1ec5-722">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-722">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="b1ec5-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ec5-726">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-726">Parameters:</span></span>

| <span data-ttu-id="b1ec5-727">Nome</span><span class="sxs-lookup"><span data-stu-id="b1ec5-727">Name</span></span> | <span data-ttu-id="b1ec5-728">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-728">Type</span></span> | <span data-ttu-id="b1ec5-729">Atributos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-729">Attributes</span></span> | <span data-ttu-id="b1ec5-730">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-730">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="b1ec5-731">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b1ec5-731">String &#124; Object</span></span>| |<span data-ttu-id="b1ec5-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b1ec5-734">**OU**</span><span class="sxs-lookup"><span data-stu-id="b1ec5-734">**OR**</span></span><br/><span data-ttu-id="b1ec5-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b1ec5-737">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-737">String</span></span> | <span data-ttu-id="b1ec5-738">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-738">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ec5-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b1ec5-741">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-741">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b1ec5-742">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-742">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ec5-743">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-743">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b1ec5-744">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-744">String</span></span> | | <span data-ttu-id="b1ec5-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b1ec5-747">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-747">String</span></span> | | <span data-ttu-id="b1ec5-748">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-748">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b1ec5-749">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-749">String</span></span> | | <span data-ttu-id="b1ec5-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="b1ec5-752">Booliano</span><span class="sxs-lookup"><span data-stu-id="b1ec5-752">Boolean</span></span> | | <span data-ttu-id="b1ec5-p144">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b1ec5-755">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-755">String</span></span> | | <span data-ttu-id="b1ec5-p145">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b1ec5-759">function</span><span class="sxs-lookup"><span data-stu-id="b1ec5-759">function</span></span> | <span data-ttu-id="b1ec5-760">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-760">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ec5-761">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-761">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b1ec5-762">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-762">Requirements</span></span>

|<span data-ttu-id="b1ec5-763">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-763">Requirement</span></span>| <span data-ttu-id="b1ec5-764">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-764">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-765">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-765">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-766">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-766">1.0</span></span>|
|[<span data-ttu-id="b1ec5-767">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-767">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-768">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-768">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-769">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-769">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-770">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-770">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b1ec5-771">Exemplos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-771">Examples</span></span>

<span data-ttu-id="b1ec5-772">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-772">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="b1ec5-773">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-773">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="b1ec5-774">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-774">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b1ec5-775">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-775">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="b1ec5-776">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-776">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="b1ec5-777">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-777">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="b1ec5-778">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-778">displayReplyForm(formData)</span></span>

<span data-ttu-id="b1ec5-779">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-779">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-780">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-780">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b1ec5-781">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-781">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b1ec5-782">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-782">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="b1ec5-p146">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ec5-786">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-786">Parameters:</span></span>

| <span data-ttu-id="b1ec5-787">Nome</span><span class="sxs-lookup"><span data-stu-id="b1ec5-787">Name</span></span> | <span data-ttu-id="b1ec5-788">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-788">Type</span></span> | <span data-ttu-id="b1ec5-789">Atributos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-789">Attributes</span></span> | <span data-ttu-id="b1ec5-790">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-790">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="b1ec5-791">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b1ec5-791">String &#124; Object</span></span>| | <span data-ttu-id="b1ec5-p147">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b1ec5-794">**OU**</span><span class="sxs-lookup"><span data-stu-id="b1ec5-794">**OR**</span></span><br/><span data-ttu-id="b1ec5-p148">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b1ec5-797">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-797">String</span></span> | <span data-ttu-id="b1ec5-798">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-798">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ec5-p149">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b1ec5-801">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-801">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b1ec5-802">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-802">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ec5-803">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-803">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b1ec5-804">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-804">String</span></span> | | <span data-ttu-id="b1ec5-p150">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b1ec5-807">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-807">String</span></span> | | <span data-ttu-id="b1ec5-808">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-808">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b1ec5-809">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-809">String</span></span> | | <span data-ttu-id="b1ec5-p151">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="b1ec5-812">Booliano</span><span class="sxs-lookup"><span data-stu-id="b1ec5-812">Boolean</span></span> | | <span data-ttu-id="b1ec5-p152">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b1ec5-815">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-815">String</span></span> | | <span data-ttu-id="b1ec5-p153">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b1ec5-819">function</span><span class="sxs-lookup"><span data-stu-id="b1ec5-819">function</span></span> | <span data-ttu-id="b1ec5-820">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-820">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ec5-821">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-821">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b1ec5-822">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-822">Requirements</span></span>

|<span data-ttu-id="b1ec5-823">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-823">Requirement</span></span>| <span data-ttu-id="b1ec5-824">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-824">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-825">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-825">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-826">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-826">1.0</span></span>|
|[<span data-ttu-id="b1ec5-827">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-827">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-828">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-828">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-829">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-829">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-830">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-830">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b1ec5-831">Exemplos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-831">Examples</span></span>

<span data-ttu-id="b1ec5-832">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-832">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="b1ec5-833">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-833">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="b1ec5-834">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-834">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b1ec5-835">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-835">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="b1ec5-836">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-836">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="b1ec5-837">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-837">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="b1ec5-838">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b1ec5-838">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="b1ec5-839">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-839">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-840">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-840">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-841">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-841">Requirements</span></span>

|<span data-ttu-id="b1ec5-842">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-842">Requirement</span></span>| <span data-ttu-id="b1ec5-843">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-843">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-844">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-844">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-845">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-845">1.0</span></span>|
|[<span data-ttu-id="b1ec5-846">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-846">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-847">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-847">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-848">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-848">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-849">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-849">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b1ec5-850">Retorna:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-850">Returns:</span></span>

<span data-ttu-id="b1ec5-851">Tipo: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-851">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b1ec5-852">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-852">Example</span></span>

<span data-ttu-id="b1ec5-853">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-853">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="b1ec5-854">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b1ec5-854">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b1ec5-855">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-855">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-856">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-856">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ec5-857">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-857">Parameters:</span></span>

|<span data-ttu-id="b1ec5-858">Nome</span><span class="sxs-lookup"><span data-stu-id="b1ec5-858">Name</span></span>| <span data-ttu-id="b1ec5-859">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-859">Type</span></span>| <span data-ttu-id="b1ec5-860">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-860">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="b1ec5-861">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="b1ec5-861">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="b1ec5-862">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-862">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ec5-863">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-863">Requirements</span></span>

|<span data-ttu-id="b1ec5-864">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-864">Requirement</span></span>| <span data-ttu-id="b1ec5-865">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-866">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-867">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-867">1.0</span></span>|
|[<span data-ttu-id="b1ec5-868">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-869">Restrito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-869">Restricted</span></span>|
|[<span data-ttu-id="b1ec5-870">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-871">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b1ec5-872">Retorna:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-872">Returns:</span></span>

<span data-ttu-id="b1ec5-873">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-873">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="b1ec5-874">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-874">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="b1ec5-875">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-875">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="b1ec5-876">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-876">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="b1ec5-877">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="b1ec5-877">Value of `entityType`</span></span> | <span data-ttu-id="b1ec5-878">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="b1ec5-878">Type of objects in returned array</span></span> | <span data-ttu-id="b1ec5-879">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="b1ec5-879">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="b1ec5-880">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-880">String</span></span> | <span data-ttu-id="b1ec5-881">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="b1ec5-881">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="b1ec5-882">Contato</span><span class="sxs-lookup"><span data-stu-id="b1ec5-882">Contact</span></span> | <span data-ttu-id="b1ec5-883">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b1ec5-883">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="b1ec5-884">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-884">String</span></span> | <span data-ttu-id="b1ec5-885">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b1ec5-885">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="b1ec5-886">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="b1ec5-886">MeetingSuggestion</span></span> | <span data-ttu-id="b1ec5-887">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b1ec5-887">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="b1ec5-888">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="b1ec5-888">PhoneNumber</span></span> | <span data-ttu-id="b1ec5-889">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="b1ec5-889">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="b1ec5-890">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="b1ec5-890">TaskSuggestion</span></span> | <span data-ttu-id="b1ec5-891">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b1ec5-891">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="b1ec5-892">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-892">String</span></span> | <span data-ttu-id="b1ec5-893">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="b1ec5-893">**Restricted**</span></span> |

<span data-ttu-id="b1ec5-894">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b1ec5-894">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="b1ec5-895">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-895">Example</span></span>

<span data-ttu-id="b1ec5-896">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-896">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="b1ec5-897">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b1ec5-897">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b1ec5-898">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-898">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-899">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-899">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b1ec5-900">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-900">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ec5-901">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-901">Parameters:</span></span>

|<span data-ttu-id="b1ec5-902">Nome</span><span class="sxs-lookup"><span data-stu-id="b1ec5-902">Name</span></span>| <span data-ttu-id="b1ec5-903">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-903">Type</span></span>| <span data-ttu-id="b1ec5-904">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-904">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b1ec5-905">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-905">String</span></span>|<span data-ttu-id="b1ec5-906">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-906">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ec5-907">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-907">Requirements</span></span>

|<span data-ttu-id="b1ec5-908">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-908">Requirement</span></span>| <span data-ttu-id="b1ec5-909">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-909">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-910">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-910">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-911">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-911">1.0</span></span>|
|[<span data-ttu-id="b1ec5-912">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-912">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-913">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-913">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-914">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-914">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-915">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-915">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b1ec5-916">Retorna:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-916">Returns:</span></span>

<span data-ttu-id="b1ec5-p155">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="b1ec5-919">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b1ec5-919">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="b1ec5-920">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b1ec5-920">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="b1ec5-921">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-921">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-922">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-922">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b1ec5-p156">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b1ec5-926">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-926">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b1ec5-927">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-927">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="b1ec5-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-931">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-931">Requirements</span></span>

|<span data-ttu-id="b1ec5-932">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-932">Requirement</span></span>| <span data-ttu-id="b1ec5-933">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-933">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-934">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-934">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-935">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-935">1.0</span></span>|
|[<span data-ttu-id="b1ec5-936">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-936">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-937">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-937">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-938">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-938">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-939">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-939">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b1ec5-940">Retorna:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-940">Returns:</span></span>

<span data-ttu-id="b1ec5-p158">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="b1ec5-943">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="b1ec5-943">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b1ec5-944">Objeto</span><span class="sxs-lookup"><span data-stu-id="b1ec5-944">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b1ec5-945">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-945">Example</span></span>

<span data-ttu-id="b1ec5-946">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-946">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="b1ec5-947">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="b1ec5-947">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="b1ec5-948">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-948">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-949">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-949">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b1ec5-950">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-950">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="b1ec5-p159">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ec5-953">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-953">Parameters:</span></span>

|<span data-ttu-id="b1ec5-954">Nome</span><span class="sxs-lookup"><span data-stu-id="b1ec5-954">Name</span></span>| <span data-ttu-id="b1ec5-955">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-955">Type</span></span>| <span data-ttu-id="b1ec5-956">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-956">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b1ec5-957">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-957">String</span></span>|<span data-ttu-id="b1ec5-958">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-958">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ec5-959">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-959">Requirements</span></span>

|<span data-ttu-id="b1ec5-960">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-960">Requirement</span></span>| <span data-ttu-id="b1ec5-961">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-962">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-963">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-963">1.0</span></span>|
|[<span data-ttu-id="b1ec5-964">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-964">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-965">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-965">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-966">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-966">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-967">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b1ec5-968">Retorna:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-968">Returns:</span></span>

<span data-ttu-id="b1ec5-969">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-969">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="b1ec5-970">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="b1ec5-970">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b1ec5-971">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="b1ec5-971">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b1ec5-972">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-972">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="b1ec5-973">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="b1ec5-973">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="b1ec5-974">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-974">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="b1ec5-p160">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ec5-977">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-977">Parameters:</span></span>

|<span data-ttu-id="b1ec5-978">Nome</span><span class="sxs-lookup"><span data-stu-id="b1ec5-978">Name</span></span>| <span data-ttu-id="b1ec5-979">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-979">Type</span></span>| <span data-ttu-id="b1ec5-980">Atributos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-980">Attributes</span></span>| <span data-ttu-id="b1ec5-981">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-981">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="b1ec5-982">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b1ec5-982">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="b1ec5-p161">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="b1ec5-986">Objeto</span><span class="sxs-lookup"><span data-stu-id="b1ec5-986">Object</span></span>| <span data-ttu-id="b1ec5-987">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-987">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-988">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-988">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b1ec5-989">Objeto</span><span class="sxs-lookup"><span data-stu-id="b1ec5-989">Object</span></span>| <span data-ttu-id="b1ec5-990">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-990">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-991">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-991">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b1ec5-992">function</span><span class="sxs-lookup"><span data-stu-id="b1ec5-992">function</span></span>||<span data-ttu-id="b1ec5-993">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-993">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b1ec5-994">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-994">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="b1ec5-995">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-995">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ec5-996">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-996">Requirements</span></span>

|<span data-ttu-id="b1ec5-997">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-997">Requirement</span></span>| <span data-ttu-id="b1ec5-998">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-998">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-999">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-999">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-1000">1.2</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1000">1.2</span></span>|
|[<span data-ttu-id="b1ec5-1001">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1001">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-1002">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1002">ReadWriteItem</span></span>|
|[<span data-ttu-id="b1ec5-1003">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1003">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-1004">Escrever</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1004">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="b1ec5-1005">Retorna:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1005">Returns:</span></span>

<span data-ttu-id="b1ec5-1006">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1006">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="b1ec5-1007">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1007">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b1ec5-1008">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1008">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b1ec5-1009">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1009">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="b1ec5-1010">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1010">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="b1ec5-p163">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-1013">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1013">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-1014">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1014">Requirements</span></span>

|<span data-ttu-id="b1ec5-1015">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1015">Requirement</span></span>| <span data-ttu-id="b1ec5-1016">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1016">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-1017">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1017">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-1018">1.6</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1018">1.6</span></span> |
|[<span data-ttu-id="b1ec5-1019">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1019">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-1020">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1020">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-1021">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1021">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-1022">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1022">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b1ec5-1023">Retorna:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1023">Returns:</span></span>

<span data-ttu-id="b1ec5-1024">Tipo: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1024">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b1ec5-1025">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1025">Example</span></span>

<span data-ttu-id="b1ec5-1026">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1026">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="b1ec5-1027">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1027">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="b1ec5-p164">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-1030">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1030">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b1ec5-p165">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b1ec5-1034">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1034">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b1ec5-1035">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1035">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="b1ec5-p166">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ec5-1039">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1039">Requirements</span></span>

|<span data-ttu-id="b1ec5-1040">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1040">Requirement</span></span>| <span data-ttu-id="b1ec5-1041">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1041">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-1042">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1042">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-1043">1.6</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1043">1.6</span></span> |
|[<span data-ttu-id="b1ec5-1044">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1044">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-1045">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1045">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-1046">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1046">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-1047">Read</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1047">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b1ec5-1048">Retorna:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1048">Returns:</span></span>

<span data-ttu-id="b1ec5-p167">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="b1ec5-1051">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1051">Example</span></span>

<span data-ttu-id="b1ec5-1052">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1052">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="b1ec5-1053">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1053">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="b1ec5-1054">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1054">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="b1ec5-p168">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ec5-1058">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1058">Parameters:</span></span>

|<span data-ttu-id="b1ec5-1059">Nome</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1059">Name</span></span>| <span data-ttu-id="b1ec5-1060">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1060">Type</span></span>| <span data-ttu-id="b1ec5-1061">Atributos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1061">Attributes</span></span>| <span data-ttu-id="b1ec5-1062">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1062">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b1ec5-1063">function</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1063">function</span></span>||<span data-ttu-id="b1ec5-1064">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1064">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b1ec5-1065">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1065">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="b1ec5-1066">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1066">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="b1ec5-1067">Objeto</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1067">Object</span></span>| <span data-ttu-id="b1ec5-1068">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1068">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-1069">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1069">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="b1ec5-1070">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1070">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ec5-1071">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1071">Requirements</span></span>

|<span data-ttu-id="b1ec5-1072">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1072">Requirement</span></span>| <span data-ttu-id="b1ec5-1073">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1073">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-1074">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1074">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-1075">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1075">1.0</span></span>|
|[<span data-ttu-id="b1ec5-1076">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1076">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-1077">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1077">ReadItem</span></span>|
|[<span data-ttu-id="b1ec5-1078">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1078">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-1079">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1079">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-1080">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1080">Example</span></span>

<span data-ttu-id="b1ec5-p171">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="b1ec5-1084">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1084">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="b1ec5-1085">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1085">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="b1ec5-p172">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ec5-1090">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1090">Parameters:</span></span>

|<span data-ttu-id="b1ec5-1091">Nome</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1091">Name</span></span>| <span data-ttu-id="b1ec5-1092">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1092">Type</span></span>| <span data-ttu-id="b1ec5-1093">Atributos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1093">Attributes</span></span>| <span data-ttu-id="b1ec5-1094">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1094">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="b1ec5-1095">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1095">String</span></span>||<span data-ttu-id="b1ec5-1096">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1096">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="b1ec5-1097">Objeto</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1097">Object</span></span>| <span data-ttu-id="b1ec5-1098">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-1099">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b1ec5-1100">Objeto</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1100">Object</span></span>| <span data-ttu-id="b1ec5-1101">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-1102">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b1ec5-1103">function</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1103">function</span></span>| <span data-ttu-id="b1ec5-1104">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-1105">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b1ec5-1106">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b1ec5-1107">Erros</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1107">Errors</span></span>

| <span data-ttu-id="b1ec5-1108">Código de erro</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1108">Error code</span></span> | <span data-ttu-id="b1ec5-1109">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="b1ec5-1110">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b1ec5-1111">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1111">Requirements</span></span>

|<span data-ttu-id="b1ec5-1112">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1112">Requirement</span></span>| <span data-ttu-id="b1ec5-1113">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-1114">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1115">1.1</span></span>|
|[<span data-ttu-id="b1ec5-1116">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="b1ec5-1118">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-1119">Escrever</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-1120">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1120">Example</span></span>

<span data-ttu-id="b1ec5-1121">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1121">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="b1ec5-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="b1ec5-1123">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="b1ec5-p173">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-1127">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1127">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="b1ec5-1128">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="b1ec5-p175">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ec5-1132">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="b1ec5-1133">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1133">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="b1ec5-1134">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1134">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="b1ec5-1135">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ec5-1136">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1136">Parameters:</span></span>

|<span data-ttu-id="b1ec5-1137">Nome</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1137">Name</span></span>| <span data-ttu-id="b1ec5-1138">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1138">Type</span></span>| <span data-ttu-id="b1ec5-1139">Atributos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1139">Attributes</span></span>| <span data-ttu-id="b1ec5-1140">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="b1ec5-1141">Object</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1141">Object</span></span>| <span data-ttu-id="b1ec5-1142">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-1143">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b1ec5-1144">Objeto</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1144">Object</span></span>| <span data-ttu-id="b1ec5-1145">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-1146">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b1ec5-1147">function</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1147">function</span></span>||<span data-ttu-id="b1ec5-1148">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b1ec5-1149">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1149">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ec5-1150">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1150">Requirements</span></span>

|<span data-ttu-id="b1ec5-1151">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1151">Requirement</span></span>| <span data-ttu-id="b1ec5-1152">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-1153">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1154">1.3</span></span>|
|[<span data-ttu-id="b1ec5-1155">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="b1ec5-1157">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-1158">Escrever</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b1ec5-1159">Exemplos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1159">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="b1ec5-p177">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="b1ec5-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="b1ec5-1163">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="b1ec5-p178">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ec5-1167">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1167">Parameters:</span></span>

|<span data-ttu-id="b1ec5-1168">Nome</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1168">Name</span></span>| <span data-ttu-id="b1ec5-1169">Tipo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1169">Type</span></span>| <span data-ttu-id="b1ec5-1170">Atributos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1170">Attributes</span></span>| <span data-ttu-id="b1ec5-1171">Descrição</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="b1ec5-1172">String</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1172">String</span></span>||<span data-ttu-id="b1ec5-p179">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="b1ec5-1176">Objeto</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1176">Object</span></span>| <span data-ttu-id="b1ec5-1177">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-1178">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b1ec5-1179">Objeto</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1179">Object</span></span>| <span data-ttu-id="b1ec5-1180">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-1181">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="b1ec5-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="b1ec5-1183">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ec5-p180">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="b1ec5-p181">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="b1ec5-1188">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="b1ec5-1189">function</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1189">function</span></span>||<span data-ttu-id="b1ec5-1190">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b1ec5-1191">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1191">Requirements</span></span>

|<span data-ttu-id="b1ec5-1192">Requisito</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1192">Requirement</span></span>| <span data-ttu-id="b1ec5-1193">Valor</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ec5-1194">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ec5-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1195">1.2</span></span>|
|[<span data-ttu-id="b1ec5-1196">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ec5-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="b1ec5-1198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ec5-1199">Escrever</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ec5-1200">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b1ec5-1200">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
