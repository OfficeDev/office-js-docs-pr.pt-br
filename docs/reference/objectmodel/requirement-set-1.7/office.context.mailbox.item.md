---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,7
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 3c2f991137441e5e425a050eeeba146c2ed540a3
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268723"
---
# <a name="item"></a><span data-ttu-id="4ac75-102">item</span><span class="sxs-lookup"><span data-stu-id="4ac75-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="4ac75-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="4ac75-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="4ac75-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="4ac75-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-106">Requirements</span></span>

|<span data-ttu-id="4ac75-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-107">Requirement</span></span>|<span data-ttu-id="4ac75-108">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-110">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-110">1.0</span></span>|
|[<span data-ttu-id="4ac75-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="4ac75-112">Restricted</span></span>|
|[<span data-ttu-id="4ac75-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4ac75-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="4ac75-115">Members and methods</span></span>

| <span data-ttu-id="4ac75-116">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-116">Member</span></span> | <span data-ttu-id="4ac75-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4ac75-118">attachments</span><span class="sxs-lookup"><span data-stu-id="4ac75-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="4ac75-119">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-119">Member</span></span> |
| [<span data-ttu-id="4ac75-120">bcc</span><span class="sxs-lookup"><span data-stu-id="4ac75-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="4ac75-121">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-121">Member</span></span> |
| [<span data-ttu-id="4ac75-122">body</span><span class="sxs-lookup"><span data-stu-id="4ac75-122">body</span></span>](#body-body) | <span data-ttu-id="4ac75-123">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-123">Member</span></span> |
| [<span data-ttu-id="4ac75-124">cc</span><span class="sxs-lookup"><span data-stu-id="4ac75-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4ac75-125">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-125">Member</span></span> |
| [<span data-ttu-id="4ac75-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="4ac75-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="4ac75-127">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-127">Member</span></span> |
| [<span data-ttu-id="4ac75-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="4ac75-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="4ac75-129">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-129">Member</span></span> |
| [<span data-ttu-id="4ac75-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="4ac75-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="4ac75-131">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-131">Member</span></span> |
| [<span data-ttu-id="4ac75-132">end</span><span class="sxs-lookup"><span data-stu-id="4ac75-132">end</span></span>](#end-datetime) | <span data-ttu-id="4ac75-133">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-133">Member</span></span> |
| [<span data-ttu-id="4ac75-134">from</span><span class="sxs-lookup"><span data-stu-id="4ac75-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="4ac75-135">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-135">Member</span></span> |
| [<span data-ttu-id="4ac75-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="4ac75-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="4ac75-137">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-137">Member</span></span> |
| [<span data-ttu-id="4ac75-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="4ac75-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="4ac75-139">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-139">Member</span></span> |
| [<span data-ttu-id="4ac75-140">itemId</span><span class="sxs-lookup"><span data-stu-id="4ac75-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="4ac75-141">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-141">Member</span></span> |
| [<span data-ttu-id="4ac75-142">itemType</span><span class="sxs-lookup"><span data-stu-id="4ac75-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="4ac75-143">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-143">Member</span></span> |
| [<span data-ttu-id="4ac75-144">location</span><span class="sxs-lookup"><span data-stu-id="4ac75-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="4ac75-145">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-145">Member</span></span> |
| [<span data-ttu-id="4ac75-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="4ac75-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="4ac75-147">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-147">Member</span></span> |
| [<span data-ttu-id="4ac75-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="4ac75-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="4ac75-149">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-149">Member</span></span> |
| [<span data-ttu-id="4ac75-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="4ac75-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4ac75-151">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-151">Member</span></span> |
| [<span data-ttu-id="4ac75-152">organizer</span><span class="sxs-lookup"><span data-stu-id="4ac75-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="4ac75-153">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-153">Member</span></span> |
| [<span data-ttu-id="4ac75-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="4ac75-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="4ac75-155">Member</span><span class="sxs-lookup"><span data-stu-id="4ac75-155">Member</span></span> |
| [<span data-ttu-id="4ac75-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="4ac75-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4ac75-157">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-157">Member</span></span> |
| [<span data-ttu-id="4ac75-158">sender</span><span class="sxs-lookup"><span data-stu-id="4ac75-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="4ac75-159">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-159">Member</span></span> |
| [<span data-ttu-id="4ac75-160">seriesid</span><span class="sxs-lookup"><span data-stu-id="4ac75-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="4ac75-161">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-161">Member</span></span> |
| [<span data-ttu-id="4ac75-162">start</span><span class="sxs-lookup"><span data-stu-id="4ac75-162">start</span></span>](#start-datetime) | <span data-ttu-id="4ac75-163">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-163">Member</span></span> |
| [<span data-ttu-id="4ac75-164">subject</span><span class="sxs-lookup"><span data-stu-id="4ac75-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="4ac75-165">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-165">Member</span></span> |
| [<span data-ttu-id="4ac75-166">to</span><span class="sxs-lookup"><span data-stu-id="4ac75-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4ac75-167">Membro</span><span class="sxs-lookup"><span data-stu-id="4ac75-167">Member</span></span> |
| [<span data-ttu-id="4ac75-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4ac75-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="4ac75-169">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-169">Method</span></span> |
| [<span data-ttu-id="4ac75-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4ac75-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="4ac75-171">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-171">Method</span></span> |
| [<span data-ttu-id="4ac75-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4ac75-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="4ac75-173">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-173">Method</span></span> |
| [<span data-ttu-id="4ac75-174">close</span><span class="sxs-lookup"><span data-stu-id="4ac75-174">close</span></span>](#close) | <span data-ttu-id="4ac75-175">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-175">Method</span></span> |
| [<span data-ttu-id="4ac75-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="4ac75-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="4ac75-177">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-177">Method</span></span> |
| [<span data-ttu-id="4ac75-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="4ac75-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="4ac75-179">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-179">Method</span></span> |
| [<span data-ttu-id="4ac75-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="4ac75-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="4ac75-181">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-181">Method</span></span> |
| [<span data-ttu-id="4ac75-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="4ac75-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="4ac75-183">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-183">Method</span></span> |
| [<span data-ttu-id="4ac75-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="4ac75-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="4ac75-185">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-185">Method</span></span> |
| [<span data-ttu-id="4ac75-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4ac75-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="4ac75-187">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-187">Method</span></span> |
| [<span data-ttu-id="4ac75-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="4ac75-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="4ac75-189">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-189">Method</span></span> |
| [<span data-ttu-id="4ac75-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4ac75-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="4ac75-191">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-191">Method</span></span> |
| [<span data-ttu-id="4ac75-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="4ac75-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="4ac75-193">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-193">Method</span></span> |
| [<span data-ttu-id="4ac75-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4ac75-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="4ac75-195">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-195">Method</span></span> |
| [<span data-ttu-id="4ac75-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="4ac75-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="4ac75-197">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-197">Method</span></span> |
| [<span data-ttu-id="4ac75-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4ac75-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="4ac75-199">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-199">Method</span></span> |
| [<span data-ttu-id="4ac75-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4ac75-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="4ac75-201">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-201">Method</span></span> |
| [<span data-ttu-id="4ac75-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="4ac75-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="4ac75-203">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-203">Method</span></span> |
| [<span data-ttu-id="4ac75-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4ac75-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="4ac75-205">Método</span><span class="sxs-lookup"><span data-stu-id="4ac75-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="4ac75-206">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-206">Example</span></span>

<span data-ttu-id="4ac75-207">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="4ac75-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="4ac75-208">Membros</span><span class="sxs-lookup"><span data-stu-id="4ac75-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-17"></a><span data-ttu-id="4ac75-209">anexos: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="4ac75-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

<span data-ttu-id="4ac75-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-212">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="4ac75-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="4ac75-213">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="4ac75-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-214">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-214">Type</span></span>

*   <span data-ttu-id="4ac75-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="4ac75-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-216">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-216">Requirements</span></span>

|<span data-ttu-id="4ac75-217">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-217">Requirement</span></span>|<span data-ttu-id="4ac75-218">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-219">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-220">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-220">1.0</span></span>|
|[<span data-ttu-id="4ac75-221">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-222">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-224">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-225">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-225">Example</span></span>

<span data-ttu-id="4ac75-226">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="4ac75-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="4ac75-227">CCO: [destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-228">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4ac75-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="4ac75-229">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="4ac75-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-230">Type</span></span>

*   [<span data-ttu-id="4ac75-231">Destinatários</span><span class="sxs-lookup"><span data-stu-id="4ac75-231">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="4ac75-232">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-232">Requirements</span></span>

|<span data-ttu-id="4ac75-233">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-233">Requirement</span></span>|<span data-ttu-id="4ac75-234">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-235">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-236">1.1</span><span class="sxs-lookup"><span data-stu-id="4ac75-236">1.1</span></span>|
|[<span data-ttu-id="4ac75-237">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-238">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-239">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-240">Escrever</span><span class="sxs-lookup"><span data-stu-id="4ac75-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-241">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-241">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-17"></a><span data-ttu-id="4ac75-242">corpo: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-242">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-243">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="4ac75-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-244">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-244">Type</span></span>

*   [<span data-ttu-id="4ac75-245">Body</span><span class="sxs-lookup"><span data-stu-id="4ac75-245">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="4ac75-246">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-246">Requirements</span></span>

|<span data-ttu-id="4ac75-247">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-247">Requirement</span></span>|<span data-ttu-id="4ac75-248">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-249">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-250">1.1</span><span class="sxs-lookup"><span data-stu-id="4ac75-250">1.1</span></span>|
|[<span data-ttu-id="4ac75-251">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-251">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-252">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-253">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-253">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-254">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-254">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-255">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-255">Example</span></span>

<span data-ttu-id="4ac75-256">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="4ac75-256">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="4ac75-257">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-257">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="4ac75-258">[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7) [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|CC: Array. <</span><span class="sxs-lookup"><span data-stu-id="4ac75-258">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-259">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4ac75-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="4ac75-260">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="4ac75-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4ac75-261">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4ac75-261">Read mode</span></span>

<span data-ttu-id="4ac75-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="4ac75-264">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4ac75-264">Compose mode</span></span>

<span data-ttu-id="4ac75-265">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="4ac75-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4ac75-266">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-266">Type</span></span>

*   <span data-ttu-id="4ac75-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-268">Requirements</span></span>

|<span data-ttu-id="4ac75-269">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-269">Requirement</span></span>|<span data-ttu-id="4ac75-270">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-271">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-272">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-272">1.0</span></span>|
|[<span data-ttu-id="4ac75-273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-274">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-275">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-276">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-276">Compose or Read</span></span>|

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="4ac75-277">(Nullable) Conversation: String</span><span class="sxs-lookup"><span data-stu-id="4ac75-277">(nullable) conversationId: String</span></span>

<span data-ttu-id="4ac75-278">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="4ac75-278">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="4ac75-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="4ac75-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-283">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-283">Type</span></span>

*   <span data-ttu-id="4ac75-284">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-284">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-285">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-285">Requirements</span></span>

|<span data-ttu-id="4ac75-286">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-286">Requirement</span></span>|<span data-ttu-id="4ac75-287">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-288">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-289">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-289">1.0</span></span>|
|[<span data-ttu-id="4ac75-290">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-291">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-292">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-293">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-294">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-294">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="4ac75-295">dateTimeCreated: data</span><span class="sxs-lookup"><span data-stu-id="4ac75-295">dateTimeCreated: Date</span></span>

<span data-ttu-id="4ac75-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-298">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-298">Type</span></span>

*   <span data-ttu-id="4ac75-299">Data</span><span class="sxs-lookup"><span data-stu-id="4ac75-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-300">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-300">Requirements</span></span>

|<span data-ttu-id="4ac75-301">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-301">Requirement</span></span>|<span data-ttu-id="4ac75-302">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-303">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-304">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-304">1.0</span></span>|
|[<span data-ttu-id="4ac75-305">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-305">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-306">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-307">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-307">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-308">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-309">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-309">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="4ac75-310">dateTimeModified: data</span><span class="sxs-lookup"><span data-stu-id="4ac75-310">dateTimeModified: Date</span></span>

<span data-ttu-id="4ac75-311">Obtém a data e a hora em que um item foi alterado pela última vez.</span><span class="sxs-lookup"><span data-stu-id="4ac75-311">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="4ac75-312">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4ac75-312">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-313">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4ac75-313">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-314">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-314">Type</span></span>

*   <span data-ttu-id="4ac75-315">Data</span><span class="sxs-lookup"><span data-stu-id="4ac75-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-316">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-316">Requirements</span></span>

|<span data-ttu-id="4ac75-317">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-317">Requirement</span></span>|<span data-ttu-id="4ac75-318">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-319">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-320">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-320">1.0</span></span>|
|[<span data-ttu-id="4ac75-321">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-321">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-322">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-323">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-323">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-324">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-325">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-325">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="4ac75-326">fim: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-326">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-327">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="4ac75-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="4ac75-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4ac75-330">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4ac75-330">Read mode</span></span>

<span data-ttu-id="4ac75-331">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-331">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="4ac75-332">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4ac75-332">Compose mode</span></span>

<span data-ttu-id="4ac75-333">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="4ac75-334">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="4ac75-334">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4ac75-335">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-335">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4ac75-336">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-336">Type</span></span>

*   <span data-ttu-id="4ac75-337">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-337">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-338">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-338">Requirements</span></span>

|<span data-ttu-id="4ac75-339">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-339">Requirement</span></span>|<span data-ttu-id="4ac75-340">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-341">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-342">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-342">1.0</span></span>|
|[<span data-ttu-id="4ac75-343">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-344">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-345">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-346">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-346">Compose or Read</span></span>|

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17fromjavascriptapioutlookofficefromviewoutlook-js-17"></a><span data-ttu-id="4ac75-347">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[de](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-347">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-348">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4ac75-348">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="4ac75-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-351">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-351">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4ac75-352">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4ac75-352">Read mode</span></span>

<span data-ttu-id="4ac75-353">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="4ac75-353">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="4ac75-354">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4ac75-354">Compose mode</span></span>

<span data-ttu-id="4ac75-355">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="4ac75-355">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4ac75-356">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-356">Type</span></span>

*   <span data-ttu-id="4ac75-357">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [de](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-357">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-358">Requirements</span></span>

|<span data-ttu-id="4ac75-359">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-359">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4ac75-360">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-361">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-361">1.0</span></span>|<span data-ttu-id="4ac75-362">1.7</span><span class="sxs-lookup"><span data-stu-id="4ac75-362">1.7</span></span>|
|[<span data-ttu-id="4ac75-363">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-364">ReadItem</span></span>|<span data-ttu-id="4ac75-365">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-365">ReadWriteItem</span></span>|
|[<span data-ttu-id="4ac75-366">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-367">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-367">Read</span></span>|<span data-ttu-id="4ac75-368">Escrever</span><span class="sxs-lookup"><span data-stu-id="4ac75-368">Compose</span></span>|

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="4ac75-369">internetMessageId: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4ac75-369">internetMessageId: String</span></span>

<span data-ttu-id="4ac75-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-372">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-372">Type</span></span>

*   <span data-ttu-id="4ac75-373">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-374">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-374">Requirements</span></span>

|<span data-ttu-id="4ac75-375">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-375">Requirement</span></span>|<span data-ttu-id="4ac75-376">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-377">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-378">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-378">1.0</span></span>|
|[<span data-ttu-id="4ac75-379">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-380">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-381">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-382">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-383">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-383">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="4ac75-384">doclass: String</span><span class="sxs-lookup"><span data-stu-id="4ac75-384">itemClass: String</span></span>

<span data-ttu-id="4ac75-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="4ac75-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="4ac75-389">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-389">Type</span></span>|<span data-ttu-id="4ac75-390">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-390">Description</span></span>|<span data-ttu-id="4ac75-391">classe de item</span><span class="sxs-lookup"><span data-stu-id="4ac75-391">item class</span></span>|
|---|---|---|
|<span data-ttu-id="4ac75-392">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="4ac75-392">Appointment items</span></span>|<span data-ttu-id="4ac75-393">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-393">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="4ac75-394">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="4ac75-394">Message items</span></span>|<span data-ttu-id="4ac75-395">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="4ac75-395">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="4ac75-396">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-396">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-397">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-397">Type</span></span>

*   <span data-ttu-id="4ac75-398">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-398">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-399">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-399">Requirements</span></span>

|<span data-ttu-id="4ac75-400">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-400">Requirement</span></span>|<span data-ttu-id="4ac75-401">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-402">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-403">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-403">1.0</span></span>|
|[<span data-ttu-id="4ac75-404">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-405">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-406">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-407">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-407">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-408">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-408">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="4ac75-409">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="4ac75-409">(nullable) itemId: String</span></span>

<span data-ttu-id="4ac75-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-412">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="4ac75-412">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4ac75-413">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4ac75-413">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="4ac75-414">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4ac75-414">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4ac75-415">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="4ac75-415">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="4ac75-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-418">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-418">Type</span></span>

*   <span data-ttu-id="4ac75-419">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-419">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-420">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-420">Requirements</span></span>

|<span data-ttu-id="4ac75-421">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-421">Requirement</span></span>|<span data-ttu-id="4ac75-422">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-423">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-424">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-424">1.0</span></span>|
|[<span data-ttu-id="4ac75-425">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-426">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-427">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-427">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-428">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-429">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-429">Example</span></span>

<span data-ttu-id="4ac75-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-17"></a><span data-ttu-id="4ac75-432">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-432">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-433">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="4ac75-433">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="4ac75-434">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="4ac75-434">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-435">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-435">Type</span></span>

*   [<span data-ttu-id="4ac75-436">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="4ac75-436">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="4ac75-437">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-437">Requirements</span></span>

|<span data-ttu-id="4ac75-438">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-438">Requirement</span></span>|<span data-ttu-id="4ac75-439">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-440">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-441">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-441">1.0</span></span>|
|[<span data-ttu-id="4ac75-442">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-443">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-444">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-445">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-445">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-446">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-446">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-17"></a><span data-ttu-id="4ac75-447">local: cadeia de caracteres | [Local](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-447">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-448">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="4ac75-448">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4ac75-449">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4ac75-449">Read mode</span></span>

<span data-ttu-id="4ac75-450">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="4ac75-450">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="4ac75-451">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4ac75-451">Compose mode</span></span>

<span data-ttu-id="4ac75-452">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="4ac75-452">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4ac75-453">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-453">Type</span></span>

*   <span data-ttu-id="4ac75-454">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-454">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-455">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-455">Requirements</span></span>

|<span data-ttu-id="4ac75-456">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-456">Requirement</span></span>|<span data-ttu-id="4ac75-457">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-458">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-459">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-459">1.0</span></span>|
|[<span data-ttu-id="4ac75-460">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-461">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-462">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-463">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-463">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="4ac75-464">normalizedSubject: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4ac75-464">normalizedSubject: String</span></span>

<span data-ttu-id="4ac75-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="4ac75-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="4ac75-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-469">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-469">Type</span></span>

*   <span data-ttu-id="4ac75-470">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-470">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-471">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-471">Requirements</span></span>

|<span data-ttu-id="4ac75-472">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-472">Requirement</span></span>|<span data-ttu-id="4ac75-473">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-474">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-475">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-475">1.0</span></span>|
|[<span data-ttu-id="4ac75-476">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-476">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-477">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-478">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-478">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-479">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-479">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-480">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-480">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-17"></a><span data-ttu-id="4ac75-481">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-481">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-482">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="4ac75-482">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-483">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-483">Type</span></span>

*   [<span data-ttu-id="4ac75-484">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="4ac75-484">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="4ac75-485">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-485">Requirements</span></span>

|<span data-ttu-id="4ac75-486">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-486">Requirement</span></span>|<span data-ttu-id="4ac75-487">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-488">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-489">1.3</span><span class="sxs-lookup"><span data-stu-id="4ac75-489">1.3</span></span>|
|[<span data-ttu-id="4ac75-490">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-491">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-492">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-493">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-493">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-494">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-494">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="4ac75-495">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.7) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="4ac75-495">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-496">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="4ac75-496">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="4ac75-497">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="4ac75-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4ac75-498">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4ac75-498">Read mode</span></span>

<span data-ttu-id="4ac75-499">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="4ac75-499">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4ac75-500">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4ac75-500">Compose mode</span></span>

<span data-ttu-id="4ac75-501">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="4ac75-501">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4ac75-502">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-502">Type</span></span>

*   <span data-ttu-id="4ac75-503">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-503">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-504">Requirements</span></span>

|<span data-ttu-id="4ac75-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-505">Requirement</span></span>|<span data-ttu-id="4ac75-506">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-507">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-508">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-508">1.0</span></span>|
|[<span data-ttu-id="4ac75-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-510">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-511">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-512">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-512">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17organizerjavascriptapioutlookofficeorganizerviewoutlook-js-17"></a><span data-ttu-id="4ac75-513">organizador: [](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4ac75-513">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-514">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-514">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4ac75-515">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4ac75-515">Read mode</span></span>

<span data-ttu-id="4ac75-516">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="4ac75-516">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="4ac75-517">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4ac75-517">Compose mode</span></span>

<span data-ttu-id="4ac75-518">A `organizer` propriedade retorna um [](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) objeto organizador que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="4ac75-518">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="4ac75-519">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-519">Type</span></span>

*   <span data-ttu-id="4ac75-520">[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4ac75-520">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-521">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-521">Requirements</span></span>

|<span data-ttu-id="4ac75-522">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-522">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4ac75-523">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-524">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-524">1.0</span></span>|<span data-ttu-id="4ac75-525">1.7</span><span class="sxs-lookup"><span data-stu-id="4ac75-525">1.7</span></span>|
|[<span data-ttu-id="4ac75-526">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-527">ReadItem</span></span>|<span data-ttu-id="4ac75-528">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-528">ReadWriteItem</span></span>|
|[<span data-ttu-id="4ac75-529">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-530">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-530">Read</span></span>|<span data-ttu-id="4ac75-531">Escrever</span><span class="sxs-lookup"><span data-stu-id="4ac75-531">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-17"></a><span data-ttu-id="4ac75-532">(anulável) recorrência [](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) : recorrência</span><span class="sxs-lookup"><span data-stu-id="4ac75-532">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-533">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="4ac75-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="4ac75-534">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="4ac75-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="4ac75-535">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="4ac75-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="4ac75-536">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="4ac75-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="4ac75-537">A `recurrence` propriedade retorna um [](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) objeto de recorrência para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="4ac75-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="4ac75-538">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="4ac75-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="4ac75-539">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="4ac75-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="4ac75-540">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="4ac75-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="4ac75-541">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="4ac75-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4ac75-542">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4ac75-542">Read mode</span></span>

<span data-ttu-id="4ac75-543">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="4ac75-543">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that represents the appointment recurrence.</span></span> <span data-ttu-id="4ac75-544">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="4ac75-544">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="4ac75-545">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4ac75-545">Compose mode</span></span>

<span data-ttu-id="4ac75-546">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="4ac75-546">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="4ac75-547">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="4ac75-547">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4ac75-548">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-548">Type</span></span>

* [<span data-ttu-id="4ac75-549">Recorrência</span><span class="sxs-lookup"><span data-stu-id="4ac75-549">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)

|<span data-ttu-id="4ac75-550">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-550">Requirement</span></span>|<span data-ttu-id="4ac75-551">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-551">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-552">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-552">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-553">1.7</span><span class="sxs-lookup"><span data-stu-id="4ac75-553">1.7</span></span>|
|[<span data-ttu-id="4ac75-554">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-554">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-555">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-555">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-556">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-556">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-557">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-557">Compose or Read</span></span>|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="4ac75-558">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.7) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="4ac75-558">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-559">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="4ac75-559">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="4ac75-560">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="4ac75-560">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4ac75-561">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4ac75-561">Read mode</span></span>

<span data-ttu-id="4ac75-562">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="4ac75-562">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4ac75-563">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4ac75-563">Compose mode</span></span>

<span data-ttu-id="4ac75-564">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="4ac75-564">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="4ac75-565">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-565">Type</span></span>

*   <span data-ttu-id="4ac75-566">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-566">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-567">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-567">Requirements</span></span>

|<span data-ttu-id="4ac75-568">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-568">Requirement</span></span>|<span data-ttu-id="4ac75-569">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-570">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-571">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-571">1.0</span></span>|
|[<span data-ttu-id="4ac75-572">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-572">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-573">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-574">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-574">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-575">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-575">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17"></a><span data-ttu-id="4ac75-576">remetente: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-576">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-p128">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="4ac75-p129">As propriedades [`from`](#from-emailaddressdetailsfrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-581">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-582">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-582">Type</span></span>

*   [<span data-ttu-id="4ac75-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4ac75-583">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="4ac75-584">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-584">Requirements</span></span>

|<span data-ttu-id="4ac75-585">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-585">Requirement</span></span>|<span data-ttu-id="4ac75-586">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-587">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-588">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-588">1.0</span></span>|
|[<span data-ttu-id="4ac75-589">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-589">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-590">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-591">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-591">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-592">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-593">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-593">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="4ac75-594">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="4ac75-594">(nullable) seriesId: String</span></span>

<span data-ttu-id="4ac75-595">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="4ac75-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="4ac75-596">No Outlook na Web e clientes de desktop, o `seriesId` retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="4ac75-596">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="4ac75-597">No entanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="4ac75-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-598">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="4ac75-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4ac75-599">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4ac75-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="4ac75-600">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4ac75-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4ac75-601">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="4ac75-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="4ac75-602">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="4ac75-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="4ac75-603">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-603">Type</span></span>

* <span data-ttu-id="4ac75-604">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-605">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-605">Requirements</span></span>

|<span data-ttu-id="4ac75-606">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-606">Requirement</span></span>|<span data-ttu-id="4ac75-607">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-608">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-609">1.7</span><span class="sxs-lookup"><span data-stu-id="4ac75-609">1.7</span></span>|
|[<span data-ttu-id="4ac75-610">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-611">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-612">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-613">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-613">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-614">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-614">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="4ac75-615">Início: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-615">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-616">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="4ac75-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="4ac75-p132">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4ac75-619">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4ac75-619">Read mode</span></span>

<span data-ttu-id="4ac75-620">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-620">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="4ac75-621">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4ac75-621">Compose mode</span></span>

<span data-ttu-id="4ac75-622">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="4ac75-623">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="4ac75-623">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4ac75-624">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-624">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4ac75-625">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-625">Type</span></span>

*   <span data-ttu-id="4ac75-626">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-626">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-627">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-627">Requirements</span></span>

|<span data-ttu-id="4ac75-628">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-628">Requirement</span></span>|<span data-ttu-id="4ac75-629">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-629">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-630">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-630">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-631">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-631">1.0</span></span>|
|[<span data-ttu-id="4ac75-632">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-632">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-633">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-633">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-634">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-634">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-635">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-635">Compose or Read</span></span>|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-17"></a><span data-ttu-id="4ac75-636">subject: cadeia de caracteres | [Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-636">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-637">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="4ac75-637">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="4ac75-638">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="4ac75-638">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4ac75-639">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4ac75-639">Read mode</span></span>

<span data-ttu-id="4ac75-p133">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="4ac75-642">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="4ac75-642">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="4ac75-643">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4ac75-643">Compose mode</span></span>

<span data-ttu-id="4ac75-644">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="4ac75-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="4ac75-645">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-645">Type</span></span>

*   <span data-ttu-id="4ac75-646">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-646">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-647">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-647">Requirements</span></span>

|<span data-ttu-id="4ac75-648">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-648">Requirement</span></span>|<span data-ttu-id="4ac75-649">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-650">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-651">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-651">1.0</span></span>|
|[<span data-ttu-id="4ac75-652">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-653">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-654">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-655">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-655">Compose or Read</span></span>|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="4ac75-656">para: Array. <[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4ac75-656">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4ac75-657">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4ac75-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="4ac75-658">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="4ac75-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4ac75-659">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="4ac75-659">Read mode</span></span>

<span data-ttu-id="4ac75-p135">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="4ac75-662">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="4ac75-662">Compose mode</span></span>

<span data-ttu-id="4ac75-663">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="4ac75-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4ac75-664">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-664">Type</span></span>

*   <span data-ttu-id="4ac75-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-666">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-666">Requirements</span></span>

|<span data-ttu-id="4ac75-667">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-667">Requirement</span></span>|<span data-ttu-id="4ac75-668">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-669">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-670">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-670">1.0</span></span>|
|[<span data-ttu-id="4ac75-671">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-672">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-673">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-674">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-674">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4ac75-675">Métodos</span><span class="sxs-lookup"><span data-stu-id="4ac75-675">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="4ac75-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4ac75-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4ac75-677">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="4ac75-677">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4ac75-678">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="4ac75-678">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="4ac75-679">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="4ac75-679">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-680">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-680">Parameters</span></span>
|<span data-ttu-id="4ac75-681">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-681">Name</span></span>|<span data-ttu-id="4ac75-682">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-682">Type</span></span>|<span data-ttu-id="4ac75-683">Atributos</span><span class="sxs-lookup"><span data-stu-id="4ac75-683">Attributes</span></span>|<span data-ttu-id="4ac75-684">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-684">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="4ac75-685">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-685">String</span></span>||<span data-ttu-id="4ac75-p136">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="4ac75-688">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-688">String</span></span>||<span data-ttu-id="4ac75-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4ac75-691">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-691">Object</span></span>|<span data-ttu-id="4ac75-692">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-692">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-693">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4ac75-693">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4ac75-694">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-694">Object</span></span>|<span data-ttu-id="4ac75-695">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-695">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-696">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-696">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="4ac75-697">Booliano</span><span class="sxs-lookup"><span data-stu-id="4ac75-697">Boolean</span></span>|<span data-ttu-id="4ac75-698">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-698">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-699">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="4ac75-699">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="4ac75-700">function</span><span class="sxs-lookup"><span data-stu-id="4ac75-700">function</span></span>|<span data-ttu-id="4ac75-701">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-701">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-702">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4ac75-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4ac75-703">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4ac75-704">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="4ac75-704">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4ac75-705">Erros</span><span class="sxs-lookup"><span data-stu-id="4ac75-705">Errors</span></span>

|<span data-ttu-id="4ac75-706">Código de erro</span><span class="sxs-lookup"><span data-stu-id="4ac75-706">Error code</span></span>|<span data-ttu-id="4ac75-707">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-707">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="4ac75-708">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="4ac75-708">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="4ac75-709">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="4ac75-709">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4ac75-710">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="4ac75-710">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-711">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-711">Requirements</span></span>

|<span data-ttu-id="4ac75-712">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-712">Requirement</span></span>|<span data-ttu-id="4ac75-713">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-713">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-714">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-714">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-715">1.1</span><span class="sxs-lookup"><span data-stu-id="4ac75-715">1.1</span></span>|
|[<span data-ttu-id="4ac75-716">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-716">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-717">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-717">ReadWriteItem</span></span>|
|[<span data-ttu-id="4ac75-718">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-718">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-719">Escrever</span><span class="sxs-lookup"><span data-stu-id="4ac75-719">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4ac75-720">Exemplos</span><span class="sxs-lookup"><span data-stu-id="4ac75-720">Examples</span></span>

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

<span data-ttu-id="4ac75-721">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="4ac75-721">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="4ac75-722">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4ac75-722">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="4ac75-723">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="4ac75-723">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="4ac75-724">Atualmente, os tipos de eventos `Office.EventType.AppointmentTimeChanged`suportados `Office.EventType.RecipientsChanged`são, e`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="4ac75-724">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-725">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-725">Parameters</span></span>

| <span data-ttu-id="4ac75-726">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-726">Name</span></span> | <span data-ttu-id="4ac75-727">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-727">Type</span></span> | <span data-ttu-id="4ac75-728">Atributos</span><span class="sxs-lookup"><span data-stu-id="4ac75-728">Attributes</span></span> | <span data-ttu-id="4ac75-729">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-729">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4ac75-730">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4ac75-730">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4ac75-731">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="4ac75-731">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="4ac75-732">Função</span><span class="sxs-lookup"><span data-stu-id="4ac75-732">Function</span></span> || <span data-ttu-id="4ac75-p138">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="4ac75-736">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-736">Object</span></span> | <span data-ttu-id="4ac75-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-737">&lt;optional&gt;</span></span> | <span data-ttu-id="4ac75-738">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4ac75-738">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4ac75-739">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-739">Object</span></span> | <span data-ttu-id="4ac75-740">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-740">&lt;optional&gt;</span></span> | <span data-ttu-id="4ac75-741">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-741">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4ac75-742">function</span><span class="sxs-lookup"><span data-stu-id="4ac75-742">function</span></span>| <span data-ttu-id="4ac75-743">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-743">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-744">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4ac75-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-745">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-745">Requirements</span></span>

|<span data-ttu-id="4ac75-746">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-746">Requirement</span></span>| <span data-ttu-id="4ac75-747">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-748">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ac75-749">1.7</span><span class="sxs-lookup"><span data-stu-id="4ac75-749">1.7</span></span> |
|[<span data-ttu-id="4ac75-750">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4ac75-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-751">ReadItem</span></span> |
|[<span data-ttu-id="4ac75-752">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4ac75-753">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-753">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="4ac75-754">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-754">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="4ac75-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4ac75-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4ac75-756">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="4ac75-756">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="4ac75-p139">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="4ac75-760">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="4ac75-760">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="4ac75-761">Se o suplemento do Office estiver em execução no Outlook na Web, o `addItemAttachmentAsync` método poderá anexar itens a itens diferentes do item que você está editando; no entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="4ac75-761">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-762">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-762">Parameters</span></span>

|<span data-ttu-id="4ac75-763">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-763">Name</span></span>|<span data-ttu-id="4ac75-764">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-764">Type</span></span>|<span data-ttu-id="4ac75-765">Atributos</span><span class="sxs-lookup"><span data-stu-id="4ac75-765">Attributes</span></span>|<span data-ttu-id="4ac75-766">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-766">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="4ac75-767">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-767">String</span></span>||<span data-ttu-id="4ac75-p140">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="4ac75-770">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4ac75-770">String</span></span>||<span data-ttu-id="4ac75-771">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="4ac75-771">The subject of the item to be attached.</span></span> <span data-ttu-id="4ac75-772">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4ac75-772">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4ac75-773">Object</span><span class="sxs-lookup"><span data-stu-id="4ac75-773">Object</span></span>|<span data-ttu-id="4ac75-774">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-774">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-775">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4ac75-775">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4ac75-776">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-776">Object</span></span>|<span data-ttu-id="4ac75-777">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-777">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-778">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-778">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4ac75-779">function</span><span class="sxs-lookup"><span data-stu-id="4ac75-779">function</span></span>|<span data-ttu-id="4ac75-780">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-780">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-781">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4ac75-781">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4ac75-782">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-782">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4ac75-783">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="4ac75-783">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4ac75-784">Erros</span><span class="sxs-lookup"><span data-stu-id="4ac75-784">Errors</span></span>

|<span data-ttu-id="4ac75-785">Código de erro</span><span class="sxs-lookup"><span data-stu-id="4ac75-785">Error code</span></span>|<span data-ttu-id="4ac75-786">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-786">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4ac75-787">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="4ac75-787">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-788">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-788">Requirements</span></span>

|<span data-ttu-id="4ac75-789">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-789">Requirement</span></span>|<span data-ttu-id="4ac75-790">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-791">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-792">1.1</span><span class="sxs-lookup"><span data-stu-id="4ac75-792">1.1</span></span>|
|[<span data-ttu-id="4ac75-793">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-794">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-794">ReadWriteItem</span></span>|
|[<span data-ttu-id="4ac75-795">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-796">Escrever</span><span class="sxs-lookup"><span data-stu-id="4ac75-796">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-797">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-797">Example</span></span>

<span data-ttu-id="4ac75-798">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-798">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="4ac75-799">close()</span><span class="sxs-lookup"><span data-stu-id="4ac75-799">close()</span></span>

<span data-ttu-id="4ac75-800">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="4ac75-800">Closes the current item that is being composed.</span></span>

<span data-ttu-id="4ac75-p142">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-803">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="4ac75-803">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="4ac75-804">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="4ac75-804">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-805">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-805">Requirements</span></span>

|<span data-ttu-id="4ac75-806">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-806">Requirement</span></span>|<span data-ttu-id="4ac75-807">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-808">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-809">1.3</span><span class="sxs-lookup"><span data-stu-id="4ac75-809">1.3</span></span>|
|[<span data-ttu-id="4ac75-810">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-811">Restrito</span><span class="sxs-lookup"><span data-stu-id="4ac75-811">Restricted</span></span>|
|[<span data-ttu-id="4ac75-812">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-813">Escrever</span><span class="sxs-lookup"><span data-stu-id="4ac75-813">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="4ac75-814">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4ac75-814">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="4ac75-815">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="4ac75-815">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-816">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4ac75-816">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4ac75-817">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="4ac75-817">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4ac75-818">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="4ac75-818">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="4ac75-819">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="4ac75-819">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="4ac75-820">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="4ac75-820">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="4ac75-821">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-821">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-822">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-822">Parameters</span></span>

|<span data-ttu-id="4ac75-823">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-823">Name</span></span>|<span data-ttu-id="4ac75-824">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-824">Type</span></span>|<span data-ttu-id="4ac75-825">Atributos</span><span class="sxs-lookup"><span data-stu-id="4ac75-825">Attributes</span></span>|<span data-ttu-id="4ac75-826">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-826">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="4ac75-827">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4ac75-827">String &#124; Object</span></span>||<span data-ttu-id="4ac75-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4ac75-830">**OU**</span><span class="sxs-lookup"><span data-stu-id="4ac75-830">**OR**</span></span><br/><span data-ttu-id="4ac75-p145">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="4ac75-833">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-833">String</span></span>|<span data-ttu-id="4ac75-834">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-834">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="4ac75-837">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-837">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="4ac75-838">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-838">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-839">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="4ac75-839">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="4ac75-840">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-840">String</span></span>||<span data-ttu-id="4ac75-p147">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="4ac75-843">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4ac75-843">String</span></span>||<span data-ttu-id="4ac75-844">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="4ac75-844">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="4ac75-845">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4ac75-845">String</span></span>||<span data-ttu-id="4ac75-p148">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="4ac75-848">Booliano</span><span class="sxs-lookup"><span data-stu-id="4ac75-848">Boolean</span></span>||<span data-ttu-id="4ac75-p149">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="4ac75-851">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-851">String</span></span>||<span data-ttu-id="4ac75-p150">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="4ac75-855">function</span><span class="sxs-lookup"><span data-stu-id="4ac75-855">function</span></span>|<span data-ttu-id="4ac75-856">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-856">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-857">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4ac75-857">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-858">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-858">Requirements</span></span>

|<span data-ttu-id="4ac75-859">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-859">Requirement</span></span>|<span data-ttu-id="4ac75-860">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-861">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-862">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-862">1.0</span></span>|
|[<span data-ttu-id="4ac75-863">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-863">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-864">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-864">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-865">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-865">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-866">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-866">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4ac75-867">Exemplos</span><span class="sxs-lookup"><span data-stu-id="4ac75-867">Examples</span></span>

<span data-ttu-id="4ac75-868">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-868">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="4ac75-869">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="4ac75-869">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="4ac75-870">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="4ac75-870">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4ac75-871">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="4ac75-871">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4ac75-872">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="4ac75-872">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4ac75-873">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-873">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="4ac75-874">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4ac75-874">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="4ac75-875">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="4ac75-875">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-876">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4ac75-876">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4ac75-877">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="4ac75-877">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4ac75-878">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="4ac75-878">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="4ac75-879">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="4ac75-879">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="4ac75-880">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="4ac75-880">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="4ac75-881">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-881">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-882">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-882">Parameters</span></span>

|<span data-ttu-id="4ac75-883">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-883">Name</span></span>|<span data-ttu-id="4ac75-884">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-884">Type</span></span>|<span data-ttu-id="4ac75-885">Atributos</span><span class="sxs-lookup"><span data-stu-id="4ac75-885">Attributes</span></span>|<span data-ttu-id="4ac75-886">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-886">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="4ac75-887">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4ac75-887">String &#124; Object</span></span>||<span data-ttu-id="4ac75-p152">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4ac75-890">**OU**</span><span class="sxs-lookup"><span data-stu-id="4ac75-890">**OR**</span></span><br/><span data-ttu-id="4ac75-p153">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="4ac75-893">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-893">String</span></span>|<span data-ttu-id="4ac75-894">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-894">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="4ac75-897">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-897">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="4ac75-898">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-898">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-899">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="4ac75-899">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="4ac75-900">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-900">String</span></span>||<span data-ttu-id="4ac75-p155">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="4ac75-903">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4ac75-903">String</span></span>||<span data-ttu-id="4ac75-904">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="4ac75-904">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="4ac75-905">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4ac75-905">String</span></span>||<span data-ttu-id="4ac75-p156">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="4ac75-908">Booliano</span><span class="sxs-lookup"><span data-stu-id="4ac75-908">Boolean</span></span>||<span data-ttu-id="4ac75-p157">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="4ac75-911">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-911">String</span></span>||<span data-ttu-id="4ac75-p158">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="4ac75-915">function</span><span class="sxs-lookup"><span data-stu-id="4ac75-915">function</span></span>|<span data-ttu-id="4ac75-916">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-916">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-917">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4ac75-917">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-918">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-918">Requirements</span></span>

|<span data-ttu-id="4ac75-919">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-919">Requirement</span></span>|<span data-ttu-id="4ac75-920">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-920">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-921">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-921">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-922">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-922">1.0</span></span>|
|[<span data-ttu-id="4ac75-923">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-923">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-924">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-924">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-925">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-925">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-926">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-926">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4ac75-927">Exemplos</span><span class="sxs-lookup"><span data-stu-id="4ac75-927">Examples</span></span>

<span data-ttu-id="4ac75-928">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-928">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="4ac75-929">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="4ac75-929">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="4ac75-930">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="4ac75-930">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4ac75-931">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="4ac75-931">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4ac75-932">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="4ac75-932">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4ac75-933">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-933">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="4ac75-934">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="4ac75-934">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="4ac75-935">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="4ac75-935">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-936">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4ac75-936">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-937">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-937">Requirements</span></span>

|<span data-ttu-id="4ac75-938">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-938">Requirement</span></span>|<span data-ttu-id="4ac75-939">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-940">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-941">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-941">1.0</span></span>|
|[<span data-ttu-id="4ac75-942">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-942">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-943">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-943">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-944">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-944">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-945">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-945">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4ac75-946">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4ac75-946">Returns:</span></span>

<span data-ttu-id="4ac75-947">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-947">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="4ac75-948">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-948">Example</span></span>

<span data-ttu-id="4ac75-949">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="4ac75-949">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="4ac75-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="4ac75-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="4ac75-951">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="4ac75-951">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-952">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4ac75-952">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-953">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-953">Parameters</span></span>

|<span data-ttu-id="4ac75-954">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-954">Name</span></span>|<span data-ttu-id="4ac75-955">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-955">Type</span></span>|<span data-ttu-id="4ac75-956">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-956">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="4ac75-957">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="4ac75-957">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.7)|<span data-ttu-id="4ac75-958">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="4ac75-958">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-959">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-959">Requirements</span></span>

|<span data-ttu-id="4ac75-960">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-960">Requirement</span></span>|<span data-ttu-id="4ac75-961">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-962">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-963">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-963">1.0</span></span>|
|[<span data-ttu-id="4ac75-964">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-964">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-965">Restrito</span><span class="sxs-lookup"><span data-stu-id="4ac75-965">Restricted</span></span>|
|[<span data-ttu-id="4ac75-966">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-966">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-967">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4ac75-968">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4ac75-968">Returns:</span></span>

<span data-ttu-id="4ac75-969">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="4ac75-969">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="4ac75-970">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="4ac75-970">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="4ac75-971">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-971">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="4ac75-972">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="4ac75-972">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="4ac75-973">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="4ac75-973">Value of `entityType`</span></span>|<span data-ttu-id="4ac75-974">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="4ac75-974">Type of objects in returned array</span></span>|<span data-ttu-id="4ac75-975">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="4ac75-975">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="4ac75-976">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-976">String</span></span>|<span data-ttu-id="4ac75-977">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="4ac75-977">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="4ac75-978">Contato</span><span class="sxs-lookup"><span data-stu-id="4ac75-978">Contact</span></span>|<span data-ttu-id="4ac75-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4ac75-979">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="4ac75-980">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-980">String</span></span>|<span data-ttu-id="4ac75-981">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4ac75-981">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="4ac75-982">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4ac75-982">MeetingSuggestion</span></span>|<span data-ttu-id="4ac75-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4ac75-983">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="4ac75-984">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4ac75-984">PhoneNumber</span></span>|<span data-ttu-id="4ac75-985">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="4ac75-985">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="4ac75-986">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4ac75-986">TaskSuggestion</span></span>|<span data-ttu-id="4ac75-987">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4ac75-987">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="4ac75-988">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-988">String</span></span>|<span data-ttu-id="4ac75-989">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="4ac75-989">**Restricted**</span></span>|

<span data-ttu-id="4ac75-990">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="4ac75-990">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

##### <a name="example"></a><span data-ttu-id="4ac75-991">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-991">Example</span></span>

<span data-ttu-id="4ac75-992">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="4ac75-992">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="4ac75-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="4ac75-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="4ac75-994">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="4ac75-994">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-995">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4ac75-995">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4ac75-996">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="4ac75-996">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-997">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-997">Parameters</span></span>

|<span data-ttu-id="4ac75-998">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-998">Name</span></span>|<span data-ttu-id="4ac75-999">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-999">Type</span></span>|<span data-ttu-id="4ac75-1000">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-1000">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="4ac75-1001">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-1001">String</span></span>|<span data-ttu-id="4ac75-1002">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1002">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-1003">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1003">Requirements</span></span>

|<span data-ttu-id="4ac75-1004">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-1004">Requirement</span></span>|<span data-ttu-id="4ac75-1005">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-1006">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-1007">1.0</span></span>|
|[<span data-ttu-id="4ac75-1008">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1008">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-1009">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-1010">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-1010">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-1011">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4ac75-1012">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4ac75-1012">Returns:</span></span>

<span data-ttu-id="4ac75-p160">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="4ac75-1015">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="4ac75-1015">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="4ac75-1016">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4ac75-1016">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="4ac75-1017">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1017">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-1018">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1018">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4ac75-p161">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4ac75-1022">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="4ac75-1022">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4ac75-1023">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1023">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4ac75-p162">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-1027">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1027">Requirements</span></span>

|<span data-ttu-id="4ac75-1028">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-1028">Requirement</span></span>|<span data-ttu-id="4ac75-1029">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-1029">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-1030">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-1030">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-1031">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-1031">1.0</span></span>|
|[<span data-ttu-id="4ac75-1032">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1032">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-1033">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-1033">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-1034">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-1034">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-1035">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-1035">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4ac75-1036">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4ac75-1036">Returns:</span></span>

<span data-ttu-id="4ac75-p163">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="4ac75-1039">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4ac75-1039">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4ac75-1040">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-1040">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4ac75-1041">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1041">Example</span></span>

<span data-ttu-id="4ac75-1042">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1042">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="4ac75-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="4ac75-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="4ac75-1044">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1044">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-1045">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1045">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4ac75-1046">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1046">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="4ac75-p164">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-1049">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-1049">Parameters</span></span>

|<span data-ttu-id="4ac75-1050">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-1050">Name</span></span>|<span data-ttu-id="4ac75-1051">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1051">Type</span></span>|<span data-ttu-id="4ac75-1052">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-1052">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="4ac75-1053">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-1053">String</span></span>|<span data-ttu-id="4ac75-1054">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1054">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-1055">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1055">Requirements</span></span>

|<span data-ttu-id="4ac75-1056">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-1056">Requirement</span></span>|<span data-ttu-id="4ac75-1057">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-1057">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-1058">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-1058">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-1059">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-1059">1.0</span></span>|
|[<span data-ttu-id="4ac75-1060">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1060">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-1061">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-1061">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-1062">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-1062">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-1063">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-1063">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4ac75-1064">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4ac75-1064">Returns:</span></span>

<span data-ttu-id="4ac75-1065">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1065">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="4ac75-1066">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4ac75-1066">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4ac75-1067">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="4ac75-1067">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4ac75-1068">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1068">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="4ac75-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="4ac75-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="4ac75-1070">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1070">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="4ac75-p165">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-1073">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-1073">Parameters</span></span>

|<span data-ttu-id="4ac75-1074">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-1074">Name</span></span>|<span data-ttu-id="4ac75-1075">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1075">Type</span></span>|<span data-ttu-id="4ac75-1076">Atributos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1076">Attributes</span></span>|<span data-ttu-id="4ac75-1077">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-1077">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="4ac75-1078">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4ac75-1078">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="4ac75-p166">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="4ac75-1082">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-1082">Object</span></span>|<span data-ttu-id="4ac75-1083">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-1084">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1084">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4ac75-1085">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-1085">Object</span></span>|<span data-ttu-id="4ac75-1086">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1086">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-1087">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1087">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4ac75-1088">function</span><span class="sxs-lookup"><span data-stu-id="4ac75-1088">function</span></span>||<span data-ttu-id="4ac75-1089">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4ac75-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4ac75-1090">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1090">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="4ac75-1091">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1091">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-1092">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1092">Requirements</span></span>

|<span data-ttu-id="4ac75-1093">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-1093">Requirement</span></span>|<span data-ttu-id="4ac75-1094">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-1094">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-1095">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-1095">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-1096">1.2</span><span class="sxs-lookup"><span data-stu-id="4ac75-1096">1.2</span></span>|
|[<span data-ttu-id="4ac75-1097">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1097">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-1098">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-1098">ReadWriteItem</span></span>|
|[<span data-ttu-id="4ac75-1099">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-1099">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-1100">Escrever</span><span class="sxs-lookup"><span data-stu-id="4ac75-1100">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4ac75-1101">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4ac75-1101">Returns:</span></span>

<span data-ttu-id="4ac75-1102">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1102">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="4ac75-1103">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4ac75-1103">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4ac75-1104">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-1104">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4ac75-1105">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1105">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="4ac75-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="4ac75-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="4ac75-1107">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1107">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="4ac75-1108">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4ac75-1108">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-1109">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1109">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-1110">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1110">Requirements</span></span>

|<span data-ttu-id="4ac75-1111">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-1111">Requirement</span></span>|<span data-ttu-id="4ac75-1112">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-1113">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-1114">1.6</span><span class="sxs-lookup"><span data-stu-id="4ac75-1114">1.6</span></span>|
|[<span data-ttu-id="4ac75-1115">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-1116">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-1116">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-1117">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-1118">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-1118">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4ac75-1119">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4ac75-1119">Returns:</span></span>

<span data-ttu-id="4ac75-1120">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4ac75-1120">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="4ac75-1121">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1121">Example</span></span>

<span data-ttu-id="4ac75-1122">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1122">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="4ac75-1123">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4ac75-1123">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="4ac75-p169">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4ac75-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-1126">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1126">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4ac75-p170">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4ac75-1130">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="4ac75-1130">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4ac75-1131">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1131">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4ac75-p171">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ac75-1135">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1135">Requirements</span></span>

|<span data-ttu-id="4ac75-1136">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-1136">Requirement</span></span>|<span data-ttu-id="4ac75-1137">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-1138">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-1139">1.6</span><span class="sxs-lookup"><span data-stu-id="4ac75-1139">1.6</span></span>|
|[<span data-ttu-id="4ac75-1140">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-1141">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-1141">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-1142">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-1143">Read</span><span class="sxs-lookup"><span data-stu-id="4ac75-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4ac75-1144">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4ac75-1144">Returns:</span></span>

<span data-ttu-id="4ac75-p172">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="4ac75-1147">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1147">Example</span></span>

<span data-ttu-id="4ac75-1148">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1148">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="4ac75-1149">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4ac75-1149">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="4ac75-1150">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1150">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="4ac75-p173">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-1154">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-1154">Parameters</span></span>

|<span data-ttu-id="4ac75-1155">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-1155">Name</span></span>|<span data-ttu-id="4ac75-1156">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1156">Type</span></span>|<span data-ttu-id="4ac75-1157">Atributos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1157">Attributes</span></span>|<span data-ttu-id="4ac75-1158">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-1158">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="4ac75-1159">function</span><span class="sxs-lookup"><span data-stu-id="4ac75-1159">function</span></span>||<span data-ttu-id="4ac75-1160">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4ac75-1160">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4ac75-1161">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1161">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4ac75-1162">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1162">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="4ac75-1163">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-1163">Object</span></span>|<span data-ttu-id="4ac75-1164">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1164">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-1165">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1165">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="4ac75-1166">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1166">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-1167">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1167">Requirements</span></span>

|<span data-ttu-id="4ac75-1168">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-1168">Requirement</span></span>|<span data-ttu-id="4ac75-1169">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-1169">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-1170">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-1170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-1171">1.0</span><span class="sxs-lookup"><span data-stu-id="4ac75-1171">1.0</span></span>|
|[<span data-ttu-id="4ac75-1172">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-1173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-1173">ReadItem</span></span>|
|[<span data-ttu-id="4ac75-1174">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-1174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-1175">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-1175">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-1176">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1176">Example</span></span>

<span data-ttu-id="4ac75-p176">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="4ac75-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4ac75-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="4ac75-1181">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1181">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="4ac75-1182">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1182">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="4ac75-1183">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1183">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="4ac75-1184">No Outlook na Web e dispositivos móveis, o identificador de anexo é válido somente dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1184">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="4ac75-1185">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1185">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-1186">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-1186">Parameters</span></span>

|<span data-ttu-id="4ac75-1187">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-1187">Name</span></span>|<span data-ttu-id="4ac75-1188">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1188">Type</span></span>|<span data-ttu-id="4ac75-1189">Atributos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1189">Attributes</span></span>|<span data-ttu-id="4ac75-1190">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-1190">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="4ac75-1191">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-1191">String</span></span>||<span data-ttu-id="4ac75-1192">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1192">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="4ac75-1193">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-1193">Object</span></span>|<span data-ttu-id="4ac75-1194">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1194">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-1195">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1195">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4ac75-1196">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-1196">Object</span></span>|<span data-ttu-id="4ac75-1197">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1197">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-1198">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1198">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4ac75-1199">function</span><span class="sxs-lookup"><span data-stu-id="4ac75-1199">function</span></span>|<span data-ttu-id="4ac75-1200">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1200">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-1201">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4ac75-1201">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4ac75-1202">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1202">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4ac75-1203">Erros</span><span class="sxs-lookup"><span data-stu-id="4ac75-1203">Errors</span></span>

|<span data-ttu-id="4ac75-1204">Código de erro</span><span class="sxs-lookup"><span data-stu-id="4ac75-1204">Error code</span></span>|<span data-ttu-id="4ac75-1205">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-1205">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="4ac75-1206">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1206">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-1207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1207">Requirements</span></span>

|<span data-ttu-id="4ac75-1208">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-1208">Requirement</span></span>|<span data-ttu-id="4ac75-1209">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-1209">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-1210">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-1210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-1211">1.1</span><span class="sxs-lookup"><span data-stu-id="4ac75-1211">1.1</span></span>|
|[<span data-ttu-id="4ac75-1212">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1212">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-1213">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-1213">ReadWriteItem</span></span>|
|[<span data-ttu-id="4ac75-1214">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-1214">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-1215">Escrever</span><span class="sxs-lookup"><span data-stu-id="4ac75-1215">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-1216">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1216">Example</span></span>

<span data-ttu-id="4ac75-1217">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1217">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="4ac75-1218">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4ac75-1218">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="4ac75-1219">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1219">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="4ac75-1220">Atualmente, os tipos de eventos `Office.EventType.AppointmentTimeChanged`suportados `Office.EventType.RecipientsChanged`são, e`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="4ac75-1220">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-1221">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-1221">Parameters</span></span>

| <span data-ttu-id="4ac75-1222">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-1222">Name</span></span> | <span data-ttu-id="4ac75-1223">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1223">Type</span></span> | <span data-ttu-id="4ac75-1224">Atributos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1224">Attributes</span></span> | <span data-ttu-id="4ac75-1225">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-1225">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4ac75-1226">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4ac75-1226">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4ac75-1227">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1227">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="4ac75-1228">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-1228">Object</span></span> | <span data-ttu-id="4ac75-1229">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="4ac75-1230">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4ac75-1231">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-1231">Object</span></span> | <span data-ttu-id="4ac75-1232">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="4ac75-1233">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4ac75-1234">function</span><span class="sxs-lookup"><span data-stu-id="4ac75-1234">function</span></span>| <span data-ttu-id="4ac75-1235">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-1236">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4ac75-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-1237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1237">Requirements</span></span>

|<span data-ttu-id="4ac75-1238">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-1238">Requirement</span></span>| <span data-ttu-id="4ac75-1239">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-1240">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-1240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ac75-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="4ac75-1241">1.7</span></span> |
|[<span data-ttu-id="4ac75-1242">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1242">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4ac75-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-1243">ReadItem</span></span> |
|[<span data-ttu-id="4ac75-1244">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4ac75-1244">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4ac75-1245">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4ac75-1245">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="4ac75-1246">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1246">Example</span></span>

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

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="4ac75-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4ac75-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="4ac75-1248">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="4ac75-1249">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1249">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="4ac75-1250">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1250">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="4ac75-1251">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1251">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-1252">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="4ac75-1253">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="4ac75-p180">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="4ac75-1257">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="4ac75-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="4ac75-1258">O Outlook no Mac não dá suporte à gravação de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1258">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="4ac75-1259">O `saveAsync` método falha quando chamado de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1259">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="4ac75-1260">Consulte [não é possível salvar uma reunião como rascunho no Outlook para Mac usando a API do Office js](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1260">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="4ac75-1261">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1261">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-1262">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-1262">Parameters</span></span>

|<span data-ttu-id="4ac75-1263">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-1263">Name</span></span>|<span data-ttu-id="4ac75-1264">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1264">Type</span></span>|<span data-ttu-id="4ac75-1265">Atributos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1265">Attributes</span></span>|<span data-ttu-id="4ac75-1266">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-1266">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4ac75-1267">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-1267">Object</span></span>|<span data-ttu-id="4ac75-1268">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-1269">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4ac75-1270">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-1270">Object</span></span>|<span data-ttu-id="4ac75-1271">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-1272">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4ac75-1273">function</span><span class="sxs-lookup"><span data-stu-id="4ac75-1273">function</span></span>||<span data-ttu-id="4ac75-1274">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4ac75-1274">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4ac75-1275">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1275">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-1276">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1276">Requirements</span></span>

|<span data-ttu-id="4ac75-1277">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-1277">Requirement</span></span>|<span data-ttu-id="4ac75-1278">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-1278">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-1279">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-1279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-1280">1.3</span><span class="sxs-lookup"><span data-stu-id="4ac75-1280">1.3</span></span>|
|[<span data-ttu-id="4ac75-1281">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1281">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-1282">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-1282">ReadWriteItem</span></span>|
|[<span data-ttu-id="4ac75-1283">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-1283">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-1284">Escrever</span><span class="sxs-lookup"><span data-stu-id="4ac75-1284">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4ac75-1285">Exemplos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1285">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="4ac75-p182">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="4ac75-1288">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="4ac75-1288">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="4ac75-1289">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1289">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="4ac75-p183">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4ac75-1293">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4ac75-1293">Parameters</span></span>

|<span data-ttu-id="4ac75-1294">Nome</span><span class="sxs-lookup"><span data-stu-id="4ac75-1294">Name</span></span>|<span data-ttu-id="4ac75-1295">Tipo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1295">Type</span></span>|<span data-ttu-id="4ac75-1296">Atributos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1296">Attributes</span></span>|<span data-ttu-id="4ac75-1297">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ac75-1297">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="4ac75-1298">String</span><span class="sxs-lookup"><span data-stu-id="4ac75-1298">String</span></span>||<span data-ttu-id="4ac75-p184">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="4ac75-1302">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-1302">Object</span></span>|<span data-ttu-id="4ac75-1303">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1303">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-1304">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1304">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4ac75-1305">Objeto</span><span class="sxs-lookup"><span data-stu-id="4ac75-1305">Object</span></span>|<span data-ttu-id="4ac75-1306">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1306">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-1307">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1307">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="4ac75-1308">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4ac75-1308">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="4ac75-1309">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4ac75-1309">&lt;optional&gt;</span></span>|<span data-ttu-id="4ac75-1310">Se `text`, o estilo atual é aplicado no Outlook na Web e clientes da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1310">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="4ac75-1311">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1311">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="4ac75-1312">Se `html` e o campo oferecer suporte a HTML (o assunto não), o estilo atual será aplicado no Outlook na Web e o estilo padrão será aplicado nos clientes da área de trabalho do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1312">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="4ac75-1313">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1313">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="4ac75-1314">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="4ac75-1314">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="4ac75-1315">function</span><span class="sxs-lookup"><span data-stu-id="4ac75-1315">function</span></span>||<span data-ttu-id="4ac75-1316">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4ac75-1316">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ac75-1317">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4ac75-1317">Requirements</span></span>

|<span data-ttu-id="4ac75-1318">Requisito</span><span class="sxs-lookup"><span data-stu-id="4ac75-1318">Requirement</span></span>|<span data-ttu-id="4ac75-1319">Valor</span><span class="sxs-lookup"><span data-stu-id="4ac75-1319">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ac75-1320">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4ac75-1320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4ac75-1321">1.2</span><span class="sxs-lookup"><span data-stu-id="4ac75-1321">1.2</span></span>|
|[<span data-ttu-id="4ac75-1322">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4ac75-1323">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4ac75-1323">ReadWriteItem</span></span>|
|[<span data-ttu-id="4ac75-1324">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4ac75-1324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4ac75-1325">Escrever</span><span class="sxs-lookup"><span data-stu-id="4ac75-1325">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4ac75-1326">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4ac75-1326">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
