---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,7
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: d400765293449899eb2e26f3d87128bc88b70000
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629676"
---
# <a name="item"></a><span data-ttu-id="786f5-102">item</span><span class="sxs-lookup"><span data-stu-id="786f5-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="786f5-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="786f5-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="786f5-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="786f5-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-106">Requirements</span></span>

|<span data-ttu-id="786f5-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-107">Requirement</span></span>|<span data-ttu-id="786f5-108">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-110">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-110">1.0</span></span>|
|[<span data-ttu-id="786f5-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="786f5-112">Restricted</span></span>|
|[<span data-ttu-id="786f5-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="786f5-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="786f5-115">Members and methods</span></span>

| <span data-ttu-id="786f5-116">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-116">Member</span></span> | <span data-ttu-id="786f5-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="786f5-118">attachments</span><span class="sxs-lookup"><span data-stu-id="786f5-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="786f5-119">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-119">Member</span></span> |
| [<span data-ttu-id="786f5-120">bcc</span><span class="sxs-lookup"><span data-stu-id="786f5-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="786f5-121">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-121">Member</span></span> |
| [<span data-ttu-id="786f5-122">body</span><span class="sxs-lookup"><span data-stu-id="786f5-122">body</span></span>](#body-body) | <span data-ttu-id="786f5-123">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-123">Member</span></span> |
| [<span data-ttu-id="786f5-124">cc</span><span class="sxs-lookup"><span data-stu-id="786f5-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="786f5-125">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-125">Member</span></span> |
| [<span data-ttu-id="786f5-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="786f5-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="786f5-127">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-127">Member</span></span> |
| [<span data-ttu-id="786f5-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="786f5-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="786f5-129">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-129">Member</span></span> |
| [<span data-ttu-id="786f5-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="786f5-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="786f5-131">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-131">Member</span></span> |
| [<span data-ttu-id="786f5-132">end</span><span class="sxs-lookup"><span data-stu-id="786f5-132">end</span></span>](#end-datetime) | <span data-ttu-id="786f5-133">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-133">Member</span></span> |
| [<span data-ttu-id="786f5-134">from</span><span class="sxs-lookup"><span data-stu-id="786f5-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="786f5-135">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-135">Member</span></span> |
| [<span data-ttu-id="786f5-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="786f5-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="786f5-137">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-137">Member</span></span> |
| [<span data-ttu-id="786f5-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="786f5-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="786f5-139">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-139">Member</span></span> |
| [<span data-ttu-id="786f5-140">itemId</span><span class="sxs-lookup"><span data-stu-id="786f5-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="786f5-141">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-141">Member</span></span> |
| [<span data-ttu-id="786f5-142">itemType</span><span class="sxs-lookup"><span data-stu-id="786f5-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="786f5-143">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-143">Member</span></span> |
| [<span data-ttu-id="786f5-144">location</span><span class="sxs-lookup"><span data-stu-id="786f5-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="786f5-145">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-145">Member</span></span> |
| [<span data-ttu-id="786f5-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="786f5-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="786f5-147">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-147">Member</span></span> |
| [<span data-ttu-id="786f5-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="786f5-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="786f5-149">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-149">Member</span></span> |
| [<span data-ttu-id="786f5-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="786f5-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="786f5-151">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-151">Member</span></span> |
| [<span data-ttu-id="786f5-152">organizer</span><span class="sxs-lookup"><span data-stu-id="786f5-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="786f5-153">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-153">Member</span></span> |
| [<span data-ttu-id="786f5-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="786f5-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="786f5-155">Member</span><span class="sxs-lookup"><span data-stu-id="786f5-155">Member</span></span> |
| [<span data-ttu-id="786f5-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="786f5-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="786f5-157">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-157">Member</span></span> |
| [<span data-ttu-id="786f5-158">sender</span><span class="sxs-lookup"><span data-stu-id="786f5-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="786f5-159">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-159">Member</span></span> |
| [<span data-ttu-id="786f5-160">seriesid</span><span class="sxs-lookup"><span data-stu-id="786f5-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="786f5-161">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-161">Member</span></span> |
| [<span data-ttu-id="786f5-162">start</span><span class="sxs-lookup"><span data-stu-id="786f5-162">start</span></span>](#start-datetime) | <span data-ttu-id="786f5-163">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-163">Member</span></span> |
| [<span data-ttu-id="786f5-164">subject</span><span class="sxs-lookup"><span data-stu-id="786f5-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="786f5-165">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-165">Member</span></span> |
| [<span data-ttu-id="786f5-166">to</span><span class="sxs-lookup"><span data-stu-id="786f5-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="786f5-167">Membro</span><span class="sxs-lookup"><span data-stu-id="786f5-167">Member</span></span> |
| [<span data-ttu-id="786f5-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="786f5-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="786f5-169">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-169">Method</span></span> |
| [<span data-ttu-id="786f5-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="786f5-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="786f5-171">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-171">Method</span></span> |
| [<span data-ttu-id="786f5-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="786f5-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="786f5-173">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-173">Method</span></span> |
| [<span data-ttu-id="786f5-174">close</span><span class="sxs-lookup"><span data-stu-id="786f5-174">close</span></span>](#close) | <span data-ttu-id="786f5-175">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-175">Method</span></span> |
| [<span data-ttu-id="786f5-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="786f5-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="786f5-177">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-177">Method</span></span> |
| [<span data-ttu-id="786f5-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="786f5-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="786f5-179">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-179">Method</span></span> |
| [<span data-ttu-id="786f5-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="786f5-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="786f5-181">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-181">Method</span></span> |
| [<span data-ttu-id="786f5-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="786f5-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="786f5-183">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-183">Method</span></span> |
| [<span data-ttu-id="786f5-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="786f5-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="786f5-185">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-185">Method</span></span> |
| [<span data-ttu-id="786f5-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="786f5-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="786f5-187">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-187">Method</span></span> |
| [<span data-ttu-id="786f5-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="786f5-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="786f5-189">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-189">Method</span></span> |
| [<span data-ttu-id="786f5-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="786f5-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="786f5-191">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-191">Method</span></span> |
| [<span data-ttu-id="786f5-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="786f5-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="786f5-193">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-193">Method</span></span> |
| [<span data-ttu-id="786f5-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="786f5-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="786f5-195">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-195">Method</span></span> |
| [<span data-ttu-id="786f5-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="786f5-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="786f5-197">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-197">Method</span></span> |
| [<span data-ttu-id="786f5-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="786f5-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="786f5-199">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-199">Method</span></span> |
| [<span data-ttu-id="786f5-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="786f5-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="786f5-201">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-201">Method</span></span> |
| [<span data-ttu-id="786f5-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="786f5-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="786f5-203">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-203">Method</span></span> |
| [<span data-ttu-id="786f5-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="786f5-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="786f5-205">Método</span><span class="sxs-lookup"><span data-stu-id="786f5-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="786f5-206">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-206">Example</span></span>

<span data-ttu-id="786f5-207">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="786f5-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="786f5-208">Members</span><span class="sxs-lookup"><span data-stu-id="786f5-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-17"></a><span data-ttu-id="786f5-209">anexos: Matriz.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="786f5-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

<span data-ttu-id="786f5-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="786f5-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-212">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="786f5-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="786f5-213">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="786f5-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-214">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-214">Type</span></span>

*   <span data-ttu-id="786f5-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="786f5-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-216">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-216">Requirements</span></span>

|<span data-ttu-id="786f5-217">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-217">Requirement</span></span>|<span data-ttu-id="786f5-218">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-219">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-220">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-220">1.0</span></span>|
|[<span data-ttu-id="786f5-221">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-222">ReadItem</span></span>|
|[<span data-ttu-id="786f5-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-224">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-225">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-225">Example</span></span>

<span data-ttu-id="786f5-226">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="786f5-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="786f5-227">cco :[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-228">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="786f5-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="786f5-229">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="786f5-229">Compose mode only.</span></span>

<span data-ttu-id="786f5-230">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-230">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="786f5-231">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="786f5-231">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="786f5-232">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-232">Get 500 members maximum.</span></span>
- <span data-ttu-id="786f5-233">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="786f5-233">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-234">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-234">Type</span></span>

*   [<span data-ttu-id="786f5-235">Destinatários</span><span class="sxs-lookup"><span data-stu-id="786f5-235">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="786f5-236">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-236">Requirements</span></span>

|<span data-ttu-id="786f5-237">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-237">Requirement</span></span>|<span data-ttu-id="786f5-238">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-239">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-240">1.1</span><span class="sxs-lookup"><span data-stu-id="786f5-240">1.1</span></span>|
|[<span data-ttu-id="786f5-241">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-242">ReadItem</span></span>|
|[<span data-ttu-id="786f5-243">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-244">Escrever</span><span class="sxs-lookup"><span data-stu-id="786f5-244">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-245">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-245">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-17"></a><span data-ttu-id="786f5-246">corpo: [Corpo](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-247">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="786f5-247">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-248">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-248">Type</span></span>

*   [<span data-ttu-id="786f5-249">Body</span><span class="sxs-lookup"><span data-stu-id="786f5-249">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="786f5-250">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-250">Requirements</span></span>

|<span data-ttu-id="786f5-251">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-251">Requirement</span></span>|<span data-ttu-id="786f5-252">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-253">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-254">1.1</span><span class="sxs-lookup"><span data-stu-id="786f5-254">1.1</span></span>|
|[<span data-ttu-id="786f5-255">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-256">ReadItem</span></span>|
|[<span data-ttu-id="786f5-257">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-258">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-259">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-259">Example</span></span>

<span data-ttu-id="786f5-260">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="786f5-260">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="786f5-261">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-261">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="786f5-262">cc : Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-263">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="786f5-263">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="786f5-264">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="786f5-264">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="786f5-265">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="786f5-265">Read mode</span></span>

<span data-ttu-id="786f5-266">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="786f5-266">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="786f5-267">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-267">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="786f5-268">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-268">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="786f5-269">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="786f5-269">Compose mode</span></span>

<span data-ttu-id="786f5-270">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="786f5-270">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="786f5-271">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-271">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="786f5-272">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="786f5-272">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="786f5-273">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-273">Get 500 members maximum.</span></span>
- <span data-ttu-id="786f5-274">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="786f5-274">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="786f5-275">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-275">Type</span></span>

*   <span data-ttu-id="786f5-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-277">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-277">Requirements</span></span>

|<span data-ttu-id="786f5-278">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-278">Requirement</span></span>|<span data-ttu-id="786f5-279">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-280">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-281">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-281">1.0</span></span>|
|[<span data-ttu-id="786f5-282">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-283">ReadItem</span></span>|
|[<span data-ttu-id="786f5-284">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-285">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-285">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="786f5-286">(anulável) conversationId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="786f5-286">(nullable) conversationId: String</span></span>

<span data-ttu-id="786f5-287">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="786f5-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="786f5-p109">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="786f5-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="786f5-p110">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="786f5-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-292">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-292">Type</span></span>

*   <span data-ttu-id="786f5-293">String</span><span class="sxs-lookup"><span data-stu-id="786f5-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-294">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-294">Requirements</span></span>

|<span data-ttu-id="786f5-295">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-295">Requirement</span></span>|<span data-ttu-id="786f5-296">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-297">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-298">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-298">1.0</span></span>|
|[<span data-ttu-id="786f5-299">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-299">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-300">ReadItem</span></span>|
|[<span data-ttu-id="786f5-301">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-301">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-302">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-302">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-303">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-303">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="786f5-304">dateTimeCreated: Data</span><span class="sxs-lookup"><span data-stu-id="786f5-304">dateTimeCreated: Date</span></span>

<span data-ttu-id="786f5-p111">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="786f5-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-307">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-307">Type</span></span>

*   <span data-ttu-id="786f5-308">Data</span><span class="sxs-lookup"><span data-stu-id="786f5-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-309">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-309">Requirements</span></span>

|<span data-ttu-id="786f5-310">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-310">Requirement</span></span>|<span data-ttu-id="786f5-311">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-312">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-313">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-313">1.0</span></span>|
|[<span data-ttu-id="786f5-314">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-315">ReadItem</span></span>|
|[<span data-ttu-id="786f5-316">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-317">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-318">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-318">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="786f5-319">dateTimeModified: Data</span><span class="sxs-lookup"><span data-stu-id="786f5-319">dateTimeModified: Date</span></span>

<span data-ttu-id="786f5-p112">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="786f5-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-322">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="786f5-322">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-323">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-323">Type</span></span>

*   <span data-ttu-id="786f5-324">Data</span><span class="sxs-lookup"><span data-stu-id="786f5-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-325">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-325">Requirements</span></span>

|<span data-ttu-id="786f5-326">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-326">Requirement</span></span>|<span data-ttu-id="786f5-327">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-328">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-329">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-329">1.0</span></span>|
|[<span data-ttu-id="786f5-330">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-331">ReadItem</span></span>|
|[<span data-ttu-id="786f5-332">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-333">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-334">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-334">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="786f5-335">fim: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-336">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="786f5-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="786f5-p113">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="786f5-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="786f5-339">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="786f5-339">Read mode</span></span>

<span data-ttu-id="786f5-340">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="786f5-340">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="786f5-341">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="786f5-341">Compose mode</span></span>

<span data-ttu-id="786f5-342">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="786f5-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="786f5-343">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="786f5-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="786f5-344">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="786f5-344">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="786f5-345">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-345">Type</span></span>

*   <span data-ttu-id="786f5-346">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-347">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-347">Requirements</span></span>

|<span data-ttu-id="786f5-348">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-348">Requirement</span></span>|<span data-ttu-id="786f5-349">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-350">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-351">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-351">1.0</span></span>|
|[<span data-ttu-id="786f5-352">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-353">ReadItem</span></span>|
|[<span data-ttu-id="786f5-354">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-355">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-355">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17fromjavascriptapioutlookofficefromviewoutlook-js-17"></a><span data-ttu-id="786f5-356">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[de](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-356">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-357">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="786f5-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="786f5-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="786f5-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-360">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="786f5-360">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="786f5-361">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="786f5-361">Read mode</span></span>

<span data-ttu-id="786f5-362">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="786f5-362">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="786f5-363">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="786f5-363">Compose mode</span></span>

<span data-ttu-id="786f5-364">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="786f5-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="786f5-365">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-365">Type</span></span>

*   <span data-ttu-id="786f5-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [de](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-367">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-367">Requirements</span></span>

|<span data-ttu-id="786f5-368">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="786f5-369">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-370">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-370">1.0</span></span>|<span data-ttu-id="786f5-371">1.7</span><span class="sxs-lookup"><span data-stu-id="786f5-371">1.7</span></span>|
|[<span data-ttu-id="786f5-372">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-373">ReadItem</span></span>|<span data-ttu-id="786f5-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="786f5-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="786f5-375">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-375">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-376">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-376">Read</span></span>|<span data-ttu-id="786f5-377">Escrever</span><span class="sxs-lookup"><span data-stu-id="786f5-377">Compose</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="786f5-378">internetMessageId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="786f5-378">internetMessageId: String</span></span>

<span data-ttu-id="786f5-p115">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="786f5-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-381">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-381">Type</span></span>

*   <span data-ttu-id="786f5-382">String</span><span class="sxs-lookup"><span data-stu-id="786f5-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-383">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-383">Requirements</span></span>

|<span data-ttu-id="786f5-384">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-384">Requirement</span></span>|<span data-ttu-id="786f5-385">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-386">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-387">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-387">1.0</span></span>|
|[<span data-ttu-id="786f5-388">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-389">ReadItem</span></span>|
|[<span data-ttu-id="786f5-390">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-391">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-392">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-392">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="786f5-393">itemClass: Cadeia de caracteres </span><span class="sxs-lookup"><span data-stu-id="786f5-393">itemClass: String</span></span>

<span data-ttu-id="786f5-p116">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="786f5-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="786f5-p117">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="786f5-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="786f5-398">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-398">Type</span></span>|<span data-ttu-id="786f5-399">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-399">Description</span></span>|<span data-ttu-id="786f5-400">classe de item</span><span class="sxs-lookup"><span data-stu-id="786f5-400">item class</span></span>|
|---|---|---|
|<span data-ttu-id="786f5-401">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="786f5-401">Appointment items</span></span>|<span data-ttu-id="786f5-402">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="786f5-402">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="786f5-403">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="786f5-403">Message items</span></span>|<span data-ttu-id="786f5-404">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="786f5-404">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="786f5-405">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="786f5-405">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-406">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-406">Type</span></span>

*   <span data-ttu-id="786f5-407">String</span><span class="sxs-lookup"><span data-stu-id="786f5-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-408">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-408">Requirements</span></span>

|<span data-ttu-id="786f5-409">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-409">Requirement</span></span>|<span data-ttu-id="786f5-410">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-411">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-412">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-412">1.0</span></span>|
|[<span data-ttu-id="786f5-413">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-414">ReadItem</span></span>|
|[<span data-ttu-id="786f5-415">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-416">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-417">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-417">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="786f5-418">(anulável) itemId: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="786f5-418">(nullable) itemId: String</span></span>

<span data-ttu-id="786f5-p118">Obtém o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="786f5-p118">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-421">O identificador retornado pela propriedade `itemId` é o mesmo que o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="786f5-421">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="786f5-422">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="786f5-422">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="786f5-423">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="786f5-423">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="786f5-424">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="786f5-424">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="786f5-p120">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-p120">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-427">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-427">Type</span></span>

*   <span data-ttu-id="786f5-428">String</span><span class="sxs-lookup"><span data-stu-id="786f5-428">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-429">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-429">Requirements</span></span>

|<span data-ttu-id="786f5-430">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-430">Requirement</span></span>|<span data-ttu-id="786f5-431">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-431">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-432">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-432">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-433">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-433">1.0</span></span>|
|[<span data-ttu-id="786f5-434">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-434">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-435">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-435">ReadItem</span></span>|
|[<span data-ttu-id="786f5-436">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-436">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-437">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-437">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-438">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-438">Example</span></span>

<span data-ttu-id="786f5-p121">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="786f5-p121">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-17"></a><span data-ttu-id="786f5-441">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-442">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="786f5-442">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="786f5-443">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="786f5-443">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-444">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-444">Type</span></span>

*   [<span data-ttu-id="786f5-445">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="786f5-445">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="786f5-446">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-446">Requirements</span></span>

|<span data-ttu-id="786f5-447">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-447">Requirement</span></span>|<span data-ttu-id="786f5-448">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-449">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-450">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-450">1.0</span></span>|
|[<span data-ttu-id="786f5-451">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-451">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-452">ReadItem</span></span>|
|[<span data-ttu-id="786f5-453">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-453">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-454">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-454">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-455">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-455">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-17"></a><span data-ttu-id="786f5-456">Local: Cadeia de caracteres[Local](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-457">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="786f5-457">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="786f5-458">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="786f5-458">Read mode</span></span>

<span data-ttu-id="786f5-459">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="786f5-459">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="786f5-460">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="786f5-460">Compose mode</span></span>

<span data-ttu-id="786f5-461">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="786f5-461">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="786f5-462">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-462">Type</span></span>

*   <span data-ttu-id="786f5-463">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-464">Requirements</span></span>

|<span data-ttu-id="786f5-465">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-465">Requirement</span></span>|<span data-ttu-id="786f5-466">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-467">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-468">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-468">1.0</span></span>|
|[<span data-ttu-id="786f5-469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-470">ReadItem</span></span>|
|[<span data-ttu-id="786f5-471">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-472">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-472">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="786f5-473">normalizedSubject: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="786f5-473">normalizedSubject: String</span></span>

<span data-ttu-id="786f5-p122">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="786f5-p122">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="786f5-p123">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="786f5-p123">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-478">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-478">Type</span></span>

*   <span data-ttu-id="786f5-479">String</span><span class="sxs-lookup"><span data-stu-id="786f5-479">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-480">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-480">Requirements</span></span>

|<span data-ttu-id="786f5-481">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-481">Requirement</span></span>|<span data-ttu-id="786f5-482">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-482">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-483">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-484">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-484">1.0</span></span>|
|[<span data-ttu-id="786f5-485">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-485">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-486">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-486">ReadItem</span></span>|
|[<span data-ttu-id="786f5-487">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-487">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-488">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-488">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-489">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-489">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-17"></a><span data-ttu-id="786f5-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-491">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="786f5-491">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-492">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-492">Type</span></span>

*   [<span data-ttu-id="786f5-493">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="786f5-493">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="786f5-494">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-494">Requirements</span></span>

|<span data-ttu-id="786f5-495">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-495">Requirement</span></span>|<span data-ttu-id="786f5-496">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-497">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-498">1.3</span><span class="sxs-lookup"><span data-stu-id="786f5-498">1.3</span></span>|
|[<span data-ttu-id="786f5-499">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-499">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-500">ReadItem</span></span>|
|[<span data-ttu-id="786f5-501">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-501">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-502">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-502">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-503">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-503">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="786f5-504">optionalAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-505">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="786f5-505">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="786f5-506">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="786f5-506">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="786f5-507">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="786f5-507">Read mode</span></span>

<span data-ttu-id="786f5-508">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="786f5-508">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="786f5-509">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-509">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="786f5-510">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-510">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="786f5-511">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="786f5-511">Compose mode</span></span>

<span data-ttu-id="786f5-512">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="786f5-512">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="786f5-513">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-513">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="786f5-514">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="786f5-514">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="786f5-515">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-515">Get 500 members maximum.</span></span>
- <span data-ttu-id="786f5-516">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="786f5-516">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="786f5-517">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-517">Type</span></span>

*   <span data-ttu-id="786f5-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-519">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-519">Requirements</span></span>

|<span data-ttu-id="786f5-520">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-520">Requirement</span></span>|<span data-ttu-id="786f5-521">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-522">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-523">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-523">1.0</span></span>|
|[<span data-ttu-id="786f5-524">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-525">ReadItem</span></span>|
|[<span data-ttu-id="786f5-526">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-527">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-527">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17organizerjavascriptapioutlookofficeorganizerviewoutlook-js-17"></a><span data-ttu-id="786f5-528">organizador: [](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="786f5-528">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-529">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="786f5-529">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="786f5-530">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="786f5-530">Read mode</span></span>

<span data-ttu-id="786f5-531">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="786f5-531">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="786f5-532">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="786f5-532">Compose mode</span></span>

<span data-ttu-id="786f5-533">A `organizer` propriedade retorna um objeto [organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="786f5-533">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="786f5-534">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-534">Type</span></span>

*   <span data-ttu-id="786f5-535">[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="786f5-535">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-536">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-536">Requirements</span></span>

|<span data-ttu-id="786f5-537">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-537">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="786f5-538">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-539">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-539">1.0</span></span>|<span data-ttu-id="786f5-540">1.7</span><span class="sxs-lookup"><span data-stu-id="786f5-540">1.7</span></span>|
|[<span data-ttu-id="786f5-541">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-542">ReadItem</span></span>|<span data-ttu-id="786f5-543">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="786f5-543">ReadWriteItem</span></span>|
|[<span data-ttu-id="786f5-544">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-545">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-545">Read</span></span>|<span data-ttu-id="786f5-546">Escrever</span><span class="sxs-lookup"><span data-stu-id="786f5-546">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-17"></a><span data-ttu-id="786f5-547">(anulável) recorrência: [recorrência](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-547">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-548">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="786f5-548">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="786f5-549">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="786f5-549">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="786f5-550">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="786f5-550">Read and compose modes for appointment items.</span></span> <span data-ttu-id="786f5-551">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="786f5-551">Read mode for meeting request items.</span></span>

<span data-ttu-id="786f5-552">A `recurrence` propriedade retorna um objeto de [recorrência](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="786f5-552">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="786f5-553">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="786f5-553">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="786f5-554">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="786f5-554">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="786f5-555">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="786f5-555">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="786f5-556">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="786f5-556">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="786f5-557">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="786f5-557">Read mode</span></span>

<span data-ttu-id="786f5-558">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="786f5-558">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that represents the appointment recurrence.</span></span> <span data-ttu-id="786f5-559">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="786f5-559">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="786f5-560">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="786f5-560">Compose mode</span></span>

<span data-ttu-id="786f5-561">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="786f5-561">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="786f5-562">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="786f5-562">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="786f5-563">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-563">Type</span></span>

* [<span data-ttu-id="786f5-564">Recorrência</span><span class="sxs-lookup"><span data-stu-id="786f5-564">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)

|<span data-ttu-id="786f5-565">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-565">Requirement</span></span>|<span data-ttu-id="786f5-566">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-567">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-568">1.7</span><span class="sxs-lookup"><span data-stu-id="786f5-568">1.7</span></span>|
|[<span data-ttu-id="786f5-569">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-570">ReadItem</span></span>|
|[<span data-ttu-id="786f5-571">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-572">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-572">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="786f5-573">requiredAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-574">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="786f5-574">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="786f5-575">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="786f5-575">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="786f5-576">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="786f5-576">Read mode</span></span>

<span data-ttu-id="786f5-577">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="786f5-577">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="786f5-578">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-578">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="786f5-579">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-579">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="786f5-580">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="786f5-580">Compose mode</span></span>

<span data-ttu-id="786f5-581">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="786f5-581">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="786f5-582">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-582">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="786f5-583">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="786f5-583">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="786f5-584">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-584">Get 500 members maximum.</span></span>
- <span data-ttu-id="786f5-585">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="786f5-585">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="786f5-586">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-586">Type</span></span>

*   <span data-ttu-id="786f5-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-588">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-588">Requirements</span></span>

|<span data-ttu-id="786f5-589">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-589">Requirement</span></span>|<span data-ttu-id="786f5-590">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-590">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-591">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-591">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-592">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-592">1.0</span></span>|
|[<span data-ttu-id="786f5-593">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-594">ReadItem</span></span>|
|[<span data-ttu-id="786f5-595">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-595">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-596">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-596">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17"></a><span data-ttu-id="786f5-597">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-p134">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="786f5-p134">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="786f5-p135">As propriedades [`from`](#from-emailaddressdetailsfrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="786f5-p135">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-602">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="786f5-602">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-603">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-603">Type</span></span>

*   [<span data-ttu-id="786f5-604">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="786f5-604">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="786f5-605">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-605">Requirements</span></span>

|<span data-ttu-id="786f5-606">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-606">Requirement</span></span>|<span data-ttu-id="786f5-607">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-608">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-609">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-609">1.0</span></span>|
|[<span data-ttu-id="786f5-610">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-611">ReadItem</span></span>|
|[<span data-ttu-id="786f5-612">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-613">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-613">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-614">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-614">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="786f5-615">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="786f5-615">(nullable) seriesId: String</span></span>

<span data-ttu-id="786f5-616">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="786f5-616">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="786f5-617">No Outlook na Web e clientes de desktop, o `seriesId` retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="786f5-617">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="786f5-618">No entanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="786f5-618">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-619">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="786f5-619">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="786f5-620">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="786f5-620">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="786f5-621">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="786f5-621">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="786f5-622">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="786f5-622">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="786f5-623">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="786f5-623">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="786f5-624">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-624">Type</span></span>

* <span data-ttu-id="786f5-625">String</span><span class="sxs-lookup"><span data-stu-id="786f5-625">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-626">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-626">Requirements</span></span>

|<span data-ttu-id="786f5-627">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-627">Requirement</span></span>|<span data-ttu-id="786f5-628">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-629">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-630">1.7</span><span class="sxs-lookup"><span data-stu-id="786f5-630">1.7</span></span>|
|[<span data-ttu-id="786f5-631">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-631">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-632">ReadItem</span></span>|
|[<span data-ttu-id="786f5-633">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-633">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-634">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-634">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-635">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-635">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="786f5-636">início: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-637">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="786f5-637">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="786f5-p138">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="786f5-p138">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="786f5-640">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="786f5-640">Read mode</span></span>

<span data-ttu-id="786f5-641">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="786f5-641">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="786f5-642">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="786f5-642">Compose mode</span></span>

<span data-ttu-id="786f5-643">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="786f5-643">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="786f5-644">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="786f5-644">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="786f5-645">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="786f5-645">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="786f5-646">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-646">Type</span></span>

*   <span data-ttu-id="786f5-647">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-648">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-648">Requirements</span></span>

|<span data-ttu-id="786f5-649">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-649">Requirement</span></span>|<span data-ttu-id="786f5-650">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-651">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-652">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-652">1.0</span></span>|
|[<span data-ttu-id="786f5-653">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-653">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-654">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-654">ReadItem</span></span>|
|[<span data-ttu-id="786f5-655">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-655">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-656">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-656">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-17"></a><span data-ttu-id="786f5-657">Assunto: Cadeia de caracteres|[Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-658">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="786f5-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="786f5-659">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="786f5-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="786f5-660">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="786f5-660">Read mode</span></span>

<span data-ttu-id="786f5-p139">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="786f5-p139">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="786f5-663">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="786f5-663">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="786f5-664">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="786f5-664">Compose mode</span></span>

<span data-ttu-id="786f5-665">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="786f5-665">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="786f5-666">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-666">Type</span></span>

*   <span data-ttu-id="786f5-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-668">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-668">Requirements</span></span>

|<span data-ttu-id="786f5-669">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-669">Requirement</span></span>|<span data-ttu-id="786f5-670">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-670">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-671">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-671">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-672">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-672">1.0</span></span>|
|[<span data-ttu-id="786f5-673">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-673">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-674">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-674">ReadItem</span></span>|
|[<span data-ttu-id="786f5-675">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-675">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-676">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-676">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="786f5-677">para: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="786f5-678">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="786f5-678">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="786f5-679">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="786f5-679">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="786f5-680">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="786f5-680">Read mode</span></span>

<span data-ttu-id="786f5-681">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="786f5-681">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="786f5-682">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-682">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="786f5-683">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-683">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="786f5-684">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="786f5-684">Compose mode</span></span>

<span data-ttu-id="786f5-685">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="786f5-685">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="786f5-686">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-686">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="786f5-687">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="786f5-687">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="786f5-688">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="786f5-688">Get 500 members maximum.</span></span>
- <span data-ttu-id="786f5-689">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="786f5-689">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="786f5-690">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-690">Type</span></span>

*   <span data-ttu-id="786f5-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-692">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-692">Requirements</span></span>

|<span data-ttu-id="786f5-693">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-693">Requirement</span></span>|<span data-ttu-id="786f5-694">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-695">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-696">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-696">1.0</span></span>|
|[<span data-ttu-id="786f5-697">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-698">ReadItem</span></span>|
|[<span data-ttu-id="786f5-699">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-700">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-700">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="786f5-701">Métodos</span><span class="sxs-lookup"><span data-stu-id="786f5-701">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="786f5-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="786f5-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="786f5-703">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="786f5-703">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="786f5-704">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="786f5-704">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="786f5-705">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="786f5-705">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-706">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-706">Parameters</span></span>
|<span data-ttu-id="786f5-707">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-707">Name</span></span>|<span data-ttu-id="786f5-708">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-708">Type</span></span>|<span data-ttu-id="786f5-709">Atributos</span><span class="sxs-lookup"><span data-stu-id="786f5-709">Attributes</span></span>|<span data-ttu-id="786f5-710">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-710">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="786f5-711">String</span><span class="sxs-lookup"><span data-stu-id="786f5-711">String</span></span>||<span data-ttu-id="786f5-p143">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="786f5-p143">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="786f5-714">String</span><span class="sxs-lookup"><span data-stu-id="786f5-714">String</span></span>||<span data-ttu-id="786f5-p144">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="786f5-p144">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="786f5-717">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-717">Object</span></span>|<span data-ttu-id="786f5-718">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-718">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-719">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="786f5-719">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="786f5-720">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-720">Object</span></span>|<span data-ttu-id="786f5-721">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-721">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-722">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-722">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="786f5-723">Booliano</span><span class="sxs-lookup"><span data-stu-id="786f5-723">Boolean</span></span>|<span data-ttu-id="786f5-724">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-724">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-725">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="786f5-725">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="786f5-726">function</span><span class="sxs-lookup"><span data-stu-id="786f5-726">function</span></span>|<span data-ttu-id="786f5-727">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-727">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-728">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="786f5-728">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="786f5-729">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="786f5-729">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="786f5-730">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="786f5-730">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="786f5-731">Erros</span><span class="sxs-lookup"><span data-stu-id="786f5-731">Errors</span></span>

|<span data-ttu-id="786f5-732">Código de erro</span><span class="sxs-lookup"><span data-stu-id="786f5-732">Error code</span></span>|<span data-ttu-id="786f5-733">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-733">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="786f5-734">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="786f5-734">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="786f5-735">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="786f5-735">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="786f5-736">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="786f5-736">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-737">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-737">Requirements</span></span>

|<span data-ttu-id="786f5-738">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-738">Requirement</span></span>|<span data-ttu-id="786f5-739">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-740">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-741">1.1</span><span class="sxs-lookup"><span data-stu-id="786f5-741">1.1</span></span>|
|[<span data-ttu-id="786f5-742">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-743">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="786f5-743">ReadWriteItem</span></span>|
|[<span data-ttu-id="786f5-744">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-745">Escrever</span><span class="sxs-lookup"><span data-stu-id="786f5-745">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="786f5-746">Exemplos</span><span class="sxs-lookup"><span data-stu-id="786f5-746">Examples</span></span>

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

<span data-ttu-id="786f5-747">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="786f5-747">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="786f5-748">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="786f5-748">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="786f5-749">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="786f5-749">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="786f5-750">Atualmente, os tipos de eventos `Office.EventType.AppointmentTimeChanged`suportados `Office.EventType.RecipientsChanged`são, e`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="786f5-750">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-751">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-751">Parameters</span></span>

| <span data-ttu-id="786f5-752">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-752">Name</span></span> | <span data-ttu-id="786f5-753">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-753">Type</span></span> | <span data-ttu-id="786f5-754">Atributos</span><span class="sxs-lookup"><span data-stu-id="786f5-754">Attributes</span></span> | <span data-ttu-id="786f5-755">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-755">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="786f5-756">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="786f5-756">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="786f5-757">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="786f5-757">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="786f5-758">Função</span><span class="sxs-lookup"><span data-stu-id="786f5-758">Function</span></span> || <span data-ttu-id="786f5-p145">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="786f5-p145">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="786f5-762">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-762">Object</span></span> | <span data-ttu-id="786f5-763">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-763">&lt;optional&gt;</span></span> | <span data-ttu-id="786f5-764">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="786f5-764">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="786f5-765">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-765">Object</span></span> | <span data-ttu-id="786f5-766">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-766">&lt;optional&gt;</span></span> | <span data-ttu-id="786f5-767">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-767">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="786f5-768">function</span><span class="sxs-lookup"><span data-stu-id="786f5-768">function</span></span>| <span data-ttu-id="786f5-769">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-769">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-770">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="786f5-770">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-771">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-771">Requirements</span></span>

|<span data-ttu-id="786f5-772">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-772">Requirement</span></span>| <span data-ttu-id="786f5-773">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-773">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-774">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-774">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="786f5-775">1.7</span><span class="sxs-lookup"><span data-stu-id="786f5-775">1.7</span></span> |
|[<span data-ttu-id="786f5-776">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-776">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="786f5-777">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-777">ReadItem</span></span> |
|[<span data-ttu-id="786f5-778">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-778">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="786f5-779">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-779">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="786f5-780">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-780">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="786f5-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="786f5-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="786f5-782">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="786f5-782">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="786f5-p146">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="786f5-p146">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="786f5-786">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="786f5-786">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="786f5-787">Se o Suplemento do Office estiver em execução no Outlook na Web, o método `addItemAttachmentAsync` pode anexar itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="786f5-787">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-788">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-788">Parameters</span></span>

|<span data-ttu-id="786f5-789">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-789">Name</span></span>|<span data-ttu-id="786f5-790">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-790">Type</span></span>|<span data-ttu-id="786f5-791">Atributos</span><span class="sxs-lookup"><span data-stu-id="786f5-791">Attributes</span></span>|<span data-ttu-id="786f5-792">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-792">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="786f5-793">String</span><span class="sxs-lookup"><span data-stu-id="786f5-793">String</span></span>||<span data-ttu-id="786f5-p147">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="786f5-p147">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="786f5-796">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="786f5-796">String</span></span>||<span data-ttu-id="786f5-797">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="786f5-797">The subject of the item to be attached.</span></span> <span data-ttu-id="786f5-798">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="786f5-798">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="786f5-799">Object</span><span class="sxs-lookup"><span data-stu-id="786f5-799">Object</span></span>|<span data-ttu-id="786f5-800">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-800">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-801">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="786f5-801">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="786f5-802">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-802">Object</span></span>|<span data-ttu-id="786f5-803">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-803">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-804">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-804">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="786f5-805">function</span><span class="sxs-lookup"><span data-stu-id="786f5-805">function</span></span>|<span data-ttu-id="786f5-806">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-806">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-807">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="786f5-807">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="786f5-808">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="786f5-808">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="786f5-809">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="786f5-809">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="786f5-810">Erros</span><span class="sxs-lookup"><span data-stu-id="786f5-810">Errors</span></span>

|<span data-ttu-id="786f5-811">Código de erro</span><span class="sxs-lookup"><span data-stu-id="786f5-811">Error code</span></span>|<span data-ttu-id="786f5-812">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-812">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="786f5-813">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="786f5-813">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-814">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-814">Requirements</span></span>

|<span data-ttu-id="786f5-815">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-815">Requirement</span></span>|<span data-ttu-id="786f5-816">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-817">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-818">1.1</span><span class="sxs-lookup"><span data-stu-id="786f5-818">1.1</span></span>|
|[<span data-ttu-id="786f5-819">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-820">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="786f5-820">ReadWriteItem</span></span>|
|[<span data-ttu-id="786f5-821">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-822">Escrever</span><span class="sxs-lookup"><span data-stu-id="786f5-822">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-823">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-823">Example</span></span>

<span data-ttu-id="786f5-824">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="786f5-824">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="786f5-825">close()</span><span class="sxs-lookup"><span data-stu-id="786f5-825">close()</span></span>

<span data-ttu-id="786f5-826">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="786f5-826">Closes the current item that is being composed.</span></span>

<span data-ttu-id="786f5-p149">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="786f5-p149">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-829">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="786f5-829">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="786f5-830">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="786f5-830">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-831">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-831">Requirements</span></span>

|<span data-ttu-id="786f5-832">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-832">Requirement</span></span>|<span data-ttu-id="786f5-833">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-834">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-835">1.3</span><span class="sxs-lookup"><span data-stu-id="786f5-835">1.3</span></span>|
|[<span data-ttu-id="786f5-836">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-837">Restrito</span><span class="sxs-lookup"><span data-stu-id="786f5-837">Restricted</span></span>|
|[<span data-ttu-id="786f5-838">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-839">Escrever</span><span class="sxs-lookup"><span data-stu-id="786f5-839">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="786f5-840">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="786f5-840">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="786f5-841">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="786f5-841">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-842">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="786f5-842">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="786f5-843">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="786f5-843">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="786f5-844">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="786f5-844">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="786f5-p150">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="786f5-p150">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-848">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-848">Parameters</span></span>

|<span data-ttu-id="786f5-849">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-849">Name</span></span>|<span data-ttu-id="786f5-850">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-850">Type</span></span>|<span data-ttu-id="786f5-851">Atributos</span><span class="sxs-lookup"><span data-stu-id="786f5-851">Attributes</span></span>|<span data-ttu-id="786f5-852">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-852">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="786f5-853">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="786f5-853">String &#124; Object</span></span>||<span data-ttu-id="786f5-p151">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="786f5-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="786f5-856">**OU**</span><span class="sxs-lookup"><span data-stu-id="786f5-856">**OR**</span></span><br/><span data-ttu-id="786f5-p152">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="786f5-p152">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="786f5-859">String</span><span class="sxs-lookup"><span data-stu-id="786f5-859">String</span></span>|<span data-ttu-id="786f5-860">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-860">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-p153">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="786f5-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="786f5-863">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-863">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="786f5-864">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-864">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-865">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="786f5-865">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="786f5-866">String</span><span class="sxs-lookup"><span data-stu-id="786f5-866">String</span></span>||<span data-ttu-id="786f5-p154">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="786f5-p154">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="786f5-869">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="786f5-869">String</span></span>||<span data-ttu-id="786f5-870">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="786f5-870">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="786f5-871">String</span><span class="sxs-lookup"><span data-stu-id="786f5-871">String</span></span>||<span data-ttu-id="786f5-p155">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="786f5-p155">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="786f5-874">Booliano</span><span class="sxs-lookup"><span data-stu-id="786f5-874">Boolean</span></span>||<span data-ttu-id="786f5-p156">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="786f5-p156">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="786f5-877">String</span><span class="sxs-lookup"><span data-stu-id="786f5-877">String</span></span>||<span data-ttu-id="786f5-p157">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="786f5-p157">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="786f5-881">function</span><span class="sxs-lookup"><span data-stu-id="786f5-881">function</span></span>|<span data-ttu-id="786f5-882">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-882">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-883">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="786f5-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-884">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-884">Requirements</span></span>

|<span data-ttu-id="786f5-885">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-885">Requirement</span></span>|<span data-ttu-id="786f5-886">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-887">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-888">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-888">1.0</span></span>|
|[<span data-ttu-id="786f5-889">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-890">ReadItem</span></span>|
|[<span data-ttu-id="786f5-891">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-892">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-892">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="786f5-893">Exemplos</span><span class="sxs-lookup"><span data-stu-id="786f5-893">Examples</span></span>

<span data-ttu-id="786f5-894">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="786f5-894">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="786f5-895">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="786f5-895">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="786f5-896">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="786f5-896">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="786f5-897">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="786f5-897">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="786f5-898">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="786f5-898">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="786f5-899">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-899">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="786f5-900">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="786f5-900">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="786f5-901">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="786f5-901">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-902">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="786f5-902">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="786f5-903">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="786f5-903">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="786f5-904">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="786f5-904">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="786f5-p158">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="786f5-p158">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-908">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-908">Parameters</span></span>

|<span data-ttu-id="786f5-909">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-909">Name</span></span>|<span data-ttu-id="786f5-910">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-910">Type</span></span>|<span data-ttu-id="786f5-911">Atributos</span><span class="sxs-lookup"><span data-stu-id="786f5-911">Attributes</span></span>|<span data-ttu-id="786f5-912">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="786f5-913">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="786f5-913">String &#124; Object</span></span>||<span data-ttu-id="786f5-p159">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="786f5-p159">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="786f5-916">**OU**</span><span class="sxs-lookup"><span data-stu-id="786f5-916">**OR**</span></span><br/><span data-ttu-id="786f5-p160">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="786f5-p160">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="786f5-919">String</span><span class="sxs-lookup"><span data-stu-id="786f5-919">String</span></span>|<span data-ttu-id="786f5-920">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-920">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-p161">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="786f5-p161">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="786f5-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="786f5-924">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-924">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-925">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="786f5-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="786f5-926">String</span><span class="sxs-lookup"><span data-stu-id="786f5-926">String</span></span>||<span data-ttu-id="786f5-p162">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="786f5-p162">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="786f5-929">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="786f5-929">String</span></span>||<span data-ttu-id="786f5-930">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="786f5-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="786f5-931">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="786f5-931">String</span></span>||<span data-ttu-id="786f5-p163">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="786f5-p163">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="786f5-934">Booliano</span><span class="sxs-lookup"><span data-stu-id="786f5-934">Boolean</span></span>||<span data-ttu-id="786f5-p164">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="786f5-p164">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="786f5-937">String</span><span class="sxs-lookup"><span data-stu-id="786f5-937">String</span></span>||<span data-ttu-id="786f5-p165">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="786f5-p165">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="786f5-941">function</span><span class="sxs-lookup"><span data-stu-id="786f5-941">function</span></span>|<span data-ttu-id="786f5-942">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-942">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-943">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="786f5-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-944">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-944">Requirements</span></span>

|<span data-ttu-id="786f5-945">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-945">Requirement</span></span>|<span data-ttu-id="786f5-946">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-947">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-948">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-948">1.0</span></span>|
|[<span data-ttu-id="786f5-949">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-950">ReadItem</span></span>|
|[<span data-ttu-id="786f5-951">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-952">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="786f5-953">Exemplos</span><span class="sxs-lookup"><span data-stu-id="786f5-953">Examples</span></span>

<span data-ttu-id="786f5-954">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="786f5-954">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="786f5-955">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="786f5-955">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="786f5-956">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="786f5-956">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="786f5-957">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="786f5-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="786f5-958">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="786f5-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="786f5-959">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="786f5-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="786f5-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="786f5-961">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="786f5-961">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-962">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="786f5-962">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-963">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-963">Requirements</span></span>

|<span data-ttu-id="786f5-964">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-964">Requirement</span></span>|<span data-ttu-id="786f5-965">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-966">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-966">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-967">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-967">1.0</span></span>|
|[<span data-ttu-id="786f5-968">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-968">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-969">ReadItem</span></span>|
|[<span data-ttu-id="786f5-970">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-970">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-971">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-971">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="786f5-972">Retorna:</span><span class="sxs-lookup"><span data-stu-id="786f5-972">Returns:</span></span>

<span data-ttu-id="786f5-973">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-973">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="786f5-974">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-974">Example</span></span>

<span data-ttu-id="786f5-975">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="786f5-975">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="786f5-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="786f5-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="786f5-977">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="786f5-977">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-978">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="786f5-978">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-979">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-979">Parameters</span></span>

|<span data-ttu-id="786f5-980">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-980">Name</span></span>|<span data-ttu-id="786f5-981">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-981">Type</span></span>|<span data-ttu-id="786f5-982">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-982">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="786f5-983">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="786f5-983">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.7)|<span data-ttu-id="786f5-984">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="786f5-984">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-985">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-985">Requirements</span></span>

|<span data-ttu-id="786f5-986">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-986">Requirement</span></span>|<span data-ttu-id="786f5-987">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-987">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-988">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-988">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-989">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-989">1.0</span></span>|
|[<span data-ttu-id="786f5-990">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-990">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-991">Restrito</span><span class="sxs-lookup"><span data-stu-id="786f5-991">Restricted</span></span>|
|[<span data-ttu-id="786f5-992">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-992">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-993">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-993">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="786f5-994">Retorna:</span><span class="sxs-lookup"><span data-stu-id="786f5-994">Returns:</span></span>

<span data-ttu-id="786f5-995">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="786f5-995">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="786f5-996">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="786f5-996">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="786f5-997">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="786f5-997">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="786f5-998">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="786f5-998">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="786f5-999">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="786f5-999">Value of `entityType`</span></span>|<span data-ttu-id="786f5-1000">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="786f5-1000">Type of objects in returned array</span></span>|<span data-ttu-id="786f5-1001">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="786f5-1001">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="786f5-1002">String</span><span class="sxs-lookup"><span data-stu-id="786f5-1002">String</span></span>|<span data-ttu-id="786f5-1003">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="786f5-1003">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="786f5-1004">Contato</span><span class="sxs-lookup"><span data-stu-id="786f5-1004">Contact</span></span>|<span data-ttu-id="786f5-1005">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="786f5-1005">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="786f5-1006">String</span><span class="sxs-lookup"><span data-stu-id="786f5-1006">String</span></span>|<span data-ttu-id="786f5-1007">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="786f5-1007">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="786f5-1008">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="786f5-1008">MeetingSuggestion</span></span>|<span data-ttu-id="786f5-1009">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="786f5-1009">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="786f5-1010">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="786f5-1010">PhoneNumber</span></span>|<span data-ttu-id="786f5-1011">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="786f5-1011">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="786f5-1012">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="786f5-1012">TaskSuggestion</span></span>|<span data-ttu-id="786f5-1013">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="786f5-1013">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="786f5-1014">String</span><span class="sxs-lookup"><span data-stu-id="786f5-1014">String</span></span>|<span data-ttu-id="786f5-1015">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="786f5-1015">**Restricted**</span></span>|

<span data-ttu-id="786f5-1016">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="786f5-1016">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

##### <a name="example"></a><span data-ttu-id="786f5-1017">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-1017">Example</span></span>

<span data-ttu-id="786f5-1018">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="786f5-1018">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="786f5-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="786f5-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="786f5-1020">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="786f5-1020">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-1021">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="786f5-1021">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="786f5-1022">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="786f5-1022">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-1023">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-1023">Parameters</span></span>

|<span data-ttu-id="786f5-1024">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-1024">Name</span></span>|<span data-ttu-id="786f5-1025">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-1025">Type</span></span>|<span data-ttu-id="786f5-1026">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-1026">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="786f5-1027">String</span><span class="sxs-lookup"><span data-stu-id="786f5-1027">String</span></span>|<span data-ttu-id="786f5-1028">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="786f5-1028">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-1029">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-1029">Requirements</span></span>

|<span data-ttu-id="786f5-1030">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-1030">Requirement</span></span>|<span data-ttu-id="786f5-1031">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-1031">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-1032">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-1032">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-1033">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-1033">1.0</span></span>|
|[<span data-ttu-id="786f5-1034">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-1034">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-1035">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-1035">ReadItem</span></span>|
|[<span data-ttu-id="786f5-1036">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-1036">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-1037">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-1037">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="786f5-1038">Retorna:</span><span class="sxs-lookup"><span data-stu-id="786f5-1038">Returns:</span></span>

<span data-ttu-id="786f5-p167">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="786f5-p167">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="786f5-1041">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="786f5-1041">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="786f5-1042">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="786f5-1042">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="786f5-1043">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="786f5-1043">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-1044">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="786f5-1044">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="786f5-p168">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="786f5-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="786f5-1048">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="786f5-1048">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="786f5-1049">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="786f5-1049">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="786f5-p169">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="786f5-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-1053">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-1053">Requirements</span></span>

|<span data-ttu-id="786f5-1054">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-1054">Requirement</span></span>|<span data-ttu-id="786f5-1055">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-1056">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-1057">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-1057">1.0</span></span>|
|[<span data-ttu-id="786f5-1058">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-1058">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-1059">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-1059">ReadItem</span></span>|
|[<span data-ttu-id="786f5-1060">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-1060">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-1061">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-1061">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="786f5-1062">Retorna:</span><span class="sxs-lookup"><span data-stu-id="786f5-1062">Returns:</span></span>

<span data-ttu-id="786f5-p170">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="786f5-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="786f5-1065">Tipo: Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-1065">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="786f5-1066">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-1066">Example</span></span>

<span data-ttu-id="786f5-1067">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="786f5-1067">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="786f5-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="786f5-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="786f5-1069">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="786f5-1069">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-1070">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="786f5-1070">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="786f5-1071">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="786f5-1071">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="786f5-p171">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="786f5-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-1074">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-1074">Parameters</span></span>

|<span data-ttu-id="786f5-1075">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-1075">Name</span></span>|<span data-ttu-id="786f5-1076">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-1076">Type</span></span>|<span data-ttu-id="786f5-1077">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-1077">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="786f5-1078">String</span><span class="sxs-lookup"><span data-stu-id="786f5-1078">String</span></span>|<span data-ttu-id="786f5-1079">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="786f5-1079">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-1080">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-1080">Requirements</span></span>

|<span data-ttu-id="786f5-1081">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-1081">Requirement</span></span>|<span data-ttu-id="786f5-1082">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-1083">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-1083">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-1084">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-1084">1.0</span></span>|
|[<span data-ttu-id="786f5-1085">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-1085">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-1086">ReadItem</span></span>|
|[<span data-ttu-id="786f5-1087">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-1087">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-1088">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-1088">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="786f5-1089">Retorna:</span><span class="sxs-lookup"><span data-stu-id="786f5-1089">Returns:</span></span>

<span data-ttu-id="786f5-1090">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="786f5-1090">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="786f5-1091">Tipo: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="786f5-1091">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="786f5-1092">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-1092">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="786f5-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="786f5-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="786f5-1094">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="786f5-1094">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="786f5-p172">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna uma cadeia de caracteres vazia para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="786f5-p172">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-1097">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-1097">Parameters</span></span>

|<span data-ttu-id="786f5-1098">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-1098">Name</span></span>|<span data-ttu-id="786f5-1099">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-1099">Type</span></span>|<span data-ttu-id="786f5-1100">Atributos</span><span class="sxs-lookup"><span data-stu-id="786f5-1100">Attributes</span></span>|<span data-ttu-id="786f5-1101">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-1101">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="786f5-1102">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="786f5-1102">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="786f5-p173">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="786f5-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="786f5-1106">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-1106">Object</span></span>|<span data-ttu-id="786f5-1107">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1107">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-1108">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="786f5-1108">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="786f5-1109">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-1109">Object</span></span>|<span data-ttu-id="786f5-1110">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1110">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-1111">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-1111">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="786f5-1112">function</span><span class="sxs-lookup"><span data-stu-id="786f5-1112">function</span></span>||<span data-ttu-id="786f5-1113">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="786f5-1113">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="786f5-1114">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="786f5-1114">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="786f5-1115">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="786f5-1115">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-1116">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-1116">Requirements</span></span>

|<span data-ttu-id="786f5-1117">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-1117">Requirement</span></span>|<span data-ttu-id="786f5-1118">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-1118">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-1119">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-1119">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-1120">1.2</span><span class="sxs-lookup"><span data-stu-id="786f5-1120">1.2</span></span>|
|[<span data-ttu-id="786f5-1121">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-1121">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-1122">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-1122">ReadItem</span></span>|
|[<span data-ttu-id="786f5-1123">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-1123">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-1124">Escrever</span><span class="sxs-lookup"><span data-stu-id="786f5-1124">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="786f5-1125">Retorna:</span><span class="sxs-lookup"><span data-stu-id="786f5-1125">Returns:</span></span>

<span data-ttu-id="786f5-1126">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="786f5-1126">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="786f5-1127">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="786f5-1127">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="786f5-1128">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-1128">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="786f5-1129">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="786f5-1129">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="786f5-1130">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="786f5-1130">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="786f5-1131">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="786f5-1131">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-1132">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="786f5-1132">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-1133">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-1133">Requirements</span></span>

|<span data-ttu-id="786f5-1134">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-1134">Requirement</span></span>|<span data-ttu-id="786f5-1135">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-1135">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-1136">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-1136">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-1137">1.6</span><span class="sxs-lookup"><span data-stu-id="786f5-1137">1.6</span></span>|
|[<span data-ttu-id="786f5-1138">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-1138">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-1139">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-1139">ReadItem</span></span>|
|[<span data-ttu-id="786f5-1140">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-1140">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-1141">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-1141">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="786f5-1142">Retorna:</span><span class="sxs-lookup"><span data-stu-id="786f5-1142">Returns:</span></span>

<span data-ttu-id="786f5-1143">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="786f5-1143">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="786f5-1144">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-1144">Example</span></span>

<span data-ttu-id="786f5-1145">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="786f5-1145">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="786f5-1146">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="786f5-1146">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="786f5-p176">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="786f5-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-1149">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="786f5-1149">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="786f5-p177">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="786f5-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="786f5-1153">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="786f5-1153">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="786f5-1154">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="786f5-1154">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="786f5-p178">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="786f5-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="786f5-1158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-1158">Requirements</span></span>

|<span data-ttu-id="786f5-1159">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-1159">Requirement</span></span>|<span data-ttu-id="786f5-1160">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-1160">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-1161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-1161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-1162">1.6</span><span class="sxs-lookup"><span data-stu-id="786f5-1162">1.6</span></span>|
|[<span data-ttu-id="786f5-1163">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-1163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-1164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-1164">ReadItem</span></span>|
|[<span data-ttu-id="786f5-1165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-1165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-1166">Read</span><span class="sxs-lookup"><span data-stu-id="786f5-1166">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="786f5-1167">Retorna:</span><span class="sxs-lookup"><span data-stu-id="786f5-1167">Returns:</span></span>

<span data-ttu-id="786f5-p179">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="786f5-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="786f5-1170">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-1170">Example</span></span>

<span data-ttu-id="786f5-1171">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="786f5-1171">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="786f5-1172">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="786f5-1172">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="786f5-1173">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="786f5-1173">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="786f5-p180">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="786f5-p180">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-1177">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-1177">Parameters</span></span>

|<span data-ttu-id="786f5-1178">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-1178">Name</span></span>|<span data-ttu-id="786f5-1179">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-1179">Type</span></span>|<span data-ttu-id="786f5-1180">Atributos</span><span class="sxs-lookup"><span data-stu-id="786f5-1180">Attributes</span></span>|<span data-ttu-id="786f5-1181">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-1181">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="786f5-1182">function</span><span class="sxs-lookup"><span data-stu-id="786f5-1182">function</span></span>||<span data-ttu-id="786f5-1183">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="786f5-1183">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="786f5-1184">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="786f5-1184">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="786f5-1185">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="786f5-1185">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="786f5-1186">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-1186">Object</span></span>|<span data-ttu-id="786f5-1187">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1187">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-1188">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-1188">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="786f5-1189">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-1189">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-1190">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-1190">Requirements</span></span>

|<span data-ttu-id="786f5-1191">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-1191">Requirement</span></span>|<span data-ttu-id="786f5-1192">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-1192">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-1193">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-1193">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-1194">1.0</span><span class="sxs-lookup"><span data-stu-id="786f5-1194">1.0</span></span>|
|[<span data-ttu-id="786f5-1195">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-1195">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-1196">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-1196">ReadItem</span></span>|
|[<span data-ttu-id="786f5-1197">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-1197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-1198">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-1198">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-1199">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-1199">Example</span></span>

<span data-ttu-id="786f5-p183">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="786f5-p183">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="786f5-1203">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="786f5-1203">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="786f5-1204">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="786f5-1204">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="786f5-1205">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="786f5-1205">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="786f5-1206">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="786f5-1206">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="786f5-1207">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="786f5-1207">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="786f5-1208">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="786f5-1208">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-1209">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-1209">Parameters</span></span>

|<span data-ttu-id="786f5-1210">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-1210">Name</span></span>|<span data-ttu-id="786f5-1211">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-1211">Type</span></span>|<span data-ttu-id="786f5-1212">Atributos</span><span class="sxs-lookup"><span data-stu-id="786f5-1212">Attributes</span></span>|<span data-ttu-id="786f5-1213">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-1213">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="786f5-1214">String</span><span class="sxs-lookup"><span data-stu-id="786f5-1214">String</span></span>||<span data-ttu-id="786f5-1215">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="786f5-1215">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="786f5-1216">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-1216">Object</span></span>|<span data-ttu-id="786f5-1217">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1217">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-1218">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="786f5-1218">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="786f5-1219">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-1219">Object</span></span>|<span data-ttu-id="786f5-1220">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1220">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-1221">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-1221">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="786f5-1222">function</span><span class="sxs-lookup"><span data-stu-id="786f5-1222">function</span></span>|<span data-ttu-id="786f5-1223">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1223">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-1224">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="786f5-1224">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="786f5-1225">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="786f5-1225">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="786f5-1226">Erros</span><span class="sxs-lookup"><span data-stu-id="786f5-1226">Errors</span></span>

|<span data-ttu-id="786f5-1227">Código de erro</span><span class="sxs-lookup"><span data-stu-id="786f5-1227">Error code</span></span>|<span data-ttu-id="786f5-1228">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-1228">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="786f5-1229">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="786f5-1229">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-1230">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-1230">Requirements</span></span>

|<span data-ttu-id="786f5-1231">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-1231">Requirement</span></span>|<span data-ttu-id="786f5-1232">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-1232">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-1233">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-1233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-1234">1.1</span><span class="sxs-lookup"><span data-stu-id="786f5-1234">1.1</span></span>|
|[<span data-ttu-id="786f5-1235">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-1235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-1236">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="786f5-1236">ReadWriteItem</span></span>|
|[<span data-ttu-id="786f5-1237">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-1237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-1238">Escrever</span><span class="sxs-lookup"><span data-stu-id="786f5-1238">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-1239">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-1239">Example</span></span>

<span data-ttu-id="786f5-1240">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="786f5-1240">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="786f5-1241">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="786f5-1241">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="786f5-1242">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="786f5-1242">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="786f5-1243">Atualmente, os tipos de eventos `Office.EventType.AppointmentTimeChanged`suportados `Office.EventType.RecipientsChanged`são, e`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="786f5-1243">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-1244">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-1244">Parameters</span></span>

| <span data-ttu-id="786f5-1245">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-1245">Name</span></span> | <span data-ttu-id="786f5-1246">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-1246">Type</span></span> | <span data-ttu-id="786f5-1247">Atributos</span><span class="sxs-lookup"><span data-stu-id="786f5-1247">Attributes</span></span> | <span data-ttu-id="786f5-1248">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-1248">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="786f5-1249">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="786f5-1249">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="786f5-1250">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="786f5-1250">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="786f5-1251">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-1251">Object</span></span> | <span data-ttu-id="786f5-1252">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1252">&lt;optional&gt;</span></span> | <span data-ttu-id="786f5-1253">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="786f5-1253">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="786f5-1254">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-1254">Object</span></span> | <span data-ttu-id="786f5-1255">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1255">&lt;optional&gt;</span></span> | <span data-ttu-id="786f5-1256">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-1256">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="786f5-1257">function</span><span class="sxs-lookup"><span data-stu-id="786f5-1257">function</span></span>| <span data-ttu-id="786f5-1258">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1258">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-1259">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="786f5-1259">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-1260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-1260">Requirements</span></span>

|<span data-ttu-id="786f5-1261">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-1261">Requirement</span></span>| <span data-ttu-id="786f5-1262">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-1262">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-1263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-1263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="786f5-1264">1.7</span><span class="sxs-lookup"><span data-stu-id="786f5-1264">1.7</span></span> |
|[<span data-ttu-id="786f5-1265">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-1265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="786f5-1266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="786f5-1266">ReadItem</span></span> |
|[<span data-ttu-id="786f5-1267">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="786f5-1267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="786f5-1268">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="786f5-1268">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="786f5-1269">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-1269">Example</span></span>

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

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="786f5-1270">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="786f5-1270">saveAsync([options], callback)</span></span>

<span data-ttu-id="786f5-1271">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="786f5-1271">Asynchronously saves an item.</span></span>

<span data-ttu-id="786f5-1272">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-1272">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="786f5-1273">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="786f5-1273">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="786f5-1274">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="786f5-1274">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-1275">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="786f5-1275">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="786f5-1276">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="786f5-1276">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="786f5-p187">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="786f5-p187">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="786f5-1280">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="786f5-1280">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="786f5-1281">O Outlook no Mac não dá suporte ao salvamento de reuniões.</span><span class="sxs-lookup"><span data-stu-id="786f5-1281">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="786f5-1282">O método `saveAsync` falha quando chamado a partir de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="786f5-1282">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="786f5-1283">Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="786f5-1283">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="786f5-1284">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="786f5-1284">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-1285">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-1285">Parameters</span></span>

|<span data-ttu-id="786f5-1286">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-1286">Name</span></span>|<span data-ttu-id="786f5-1287">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-1287">Type</span></span>|<span data-ttu-id="786f5-1288">Atributos</span><span class="sxs-lookup"><span data-stu-id="786f5-1288">Attributes</span></span>|<span data-ttu-id="786f5-1289">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-1289">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="786f5-1290">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-1290">Object</span></span>|<span data-ttu-id="786f5-1291">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1291">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-1292">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="786f5-1292">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="786f5-1293">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-1293">Object</span></span>|<span data-ttu-id="786f5-1294">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1294">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-1295">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-1295">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="786f5-1296">function</span><span class="sxs-lookup"><span data-stu-id="786f5-1296">function</span></span>||<span data-ttu-id="786f5-1297">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="786f5-1297">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="786f5-1298">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="786f5-1298">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-1299">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-1299">Requirements</span></span>

|<span data-ttu-id="786f5-1300">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-1300">Requirement</span></span>|<span data-ttu-id="786f5-1301">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-1301">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-1302">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-1302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-1303">1.3</span><span class="sxs-lookup"><span data-stu-id="786f5-1303">1.3</span></span>|
|[<span data-ttu-id="786f5-1304">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-1304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-1305">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="786f5-1305">ReadWriteItem</span></span>|
|[<span data-ttu-id="786f5-1306">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-1306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-1307">Escrever</span><span class="sxs-lookup"><span data-stu-id="786f5-1307">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="786f5-1308">Exemplos</span><span class="sxs-lookup"><span data-stu-id="786f5-1308">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="786f5-p189">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="786f5-p189">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="786f5-1311">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="786f5-1311">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="786f5-1312">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="786f5-1312">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="786f5-p190">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="786f5-p190">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="786f5-1316">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="786f5-1316">Parameters</span></span>

|<span data-ttu-id="786f5-1317">Nome</span><span class="sxs-lookup"><span data-stu-id="786f5-1317">Name</span></span>|<span data-ttu-id="786f5-1318">Tipo</span><span class="sxs-lookup"><span data-stu-id="786f5-1318">Type</span></span>|<span data-ttu-id="786f5-1319">Atributos</span><span class="sxs-lookup"><span data-stu-id="786f5-1319">Attributes</span></span>|<span data-ttu-id="786f5-1320">Descrição</span><span class="sxs-lookup"><span data-stu-id="786f5-1320">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="786f5-1321">String</span><span class="sxs-lookup"><span data-stu-id="786f5-1321">String</span></span>||<span data-ttu-id="786f5-p191">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="786f5-p191">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="786f5-1325">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-1325">Object</span></span>|<span data-ttu-id="786f5-1326">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1326">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-1327">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="786f5-1327">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="786f5-1328">Objeto</span><span class="sxs-lookup"><span data-stu-id="786f5-1328">Object</span></span>|<span data-ttu-id="786f5-1329">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1329">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-1330">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="786f5-1330">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="786f5-1331">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="786f5-1331">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="786f5-1332">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="786f5-1332">&lt;optional&gt;</span></span>|<span data-ttu-id="786f5-1333">Se `text`, o estilo atual é aplicado nos clientes do Outlook na Web e do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="786f5-1333">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="786f5-1334">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="786f5-1334">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="786f5-1335">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook na Web e o estilo padrão é aplicado nos clientes do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="786f5-1335">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="786f5-1336">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="786f5-1336">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="786f5-1337">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="786f5-1337">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="786f5-1338">function</span><span class="sxs-lookup"><span data-stu-id="786f5-1338">function</span></span>||<span data-ttu-id="786f5-1339">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="786f5-1339">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="786f5-1340">Requisitos</span><span class="sxs-lookup"><span data-stu-id="786f5-1340">Requirements</span></span>

|<span data-ttu-id="786f5-1341">Requisito</span><span class="sxs-lookup"><span data-stu-id="786f5-1341">Requirement</span></span>|<span data-ttu-id="786f5-1342">Valor</span><span class="sxs-lookup"><span data-stu-id="786f5-1342">Value</span></span>|
|---|---|
|[<span data-ttu-id="786f5-1343">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="786f5-1343">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="786f5-1344">1.2</span><span class="sxs-lookup"><span data-stu-id="786f5-1344">1.2</span></span>|
|[<span data-ttu-id="786f5-1345">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="786f5-1345">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="786f5-1346">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="786f5-1346">ReadWriteItem</span></span>|
|[<span data-ttu-id="786f5-1347">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="786f5-1347">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="786f5-1348">Escrever</span><span class="sxs-lookup"><span data-stu-id="786f5-1348">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="786f5-1349">Exemplo</span><span class="sxs-lookup"><span data-stu-id="786f5-1349">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
