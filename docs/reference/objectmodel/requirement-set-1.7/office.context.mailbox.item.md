---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,7
description: ''
ms.date: 11/06/2019
localization_priority: Normal
ms.openlocfilehash: 1c0948490c5c0b77252a8605b43f85dd529f2897
ms.sourcegitcommit: 08c0b9ff319c391922fa43d3c2e9783cf6b53b1b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/08/2019
ms.locfileid: "38066211"
---
# <a name="item"></a><span data-ttu-id="78f1e-102">item</span><span class="sxs-lookup"><span data-stu-id="78f1e-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="78f1e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="78f1e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="78f1e-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="78f1e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-106">Requirements</span></span>

|<span data-ttu-id="78f1e-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-107">Requirement</span></span>|<span data-ttu-id="78f1e-108">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-110">1.0</span></span>|
|[<span data-ttu-id="78f1e-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="78f1e-112">Restricted</span></span>|
|[<span data-ttu-id="78f1e-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="78f1e-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="78f1e-115">Members and methods</span></span>

| <span data-ttu-id="78f1e-116">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-116">Member</span></span> | <span data-ttu-id="78f1e-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="78f1e-118">attachments</span><span class="sxs-lookup"><span data-stu-id="78f1e-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="78f1e-119">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-119">Member</span></span> |
| [<span data-ttu-id="78f1e-120">bcc</span><span class="sxs-lookup"><span data-stu-id="78f1e-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="78f1e-121">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-121">Member</span></span> |
| [<span data-ttu-id="78f1e-122">body</span><span class="sxs-lookup"><span data-stu-id="78f1e-122">body</span></span>](#body-body) | <span data-ttu-id="78f1e-123">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-123">Member</span></span> |
| [<span data-ttu-id="78f1e-124">cc</span><span class="sxs-lookup"><span data-stu-id="78f1e-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="78f1e-125">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-125">Member</span></span> |
| [<span data-ttu-id="78f1e-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="78f1e-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="78f1e-127">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-127">Member</span></span> |
| [<span data-ttu-id="78f1e-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="78f1e-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="78f1e-129">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-129">Member</span></span> |
| [<span data-ttu-id="78f1e-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="78f1e-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="78f1e-131">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-131">Member</span></span> |
| [<span data-ttu-id="78f1e-132">end</span><span class="sxs-lookup"><span data-stu-id="78f1e-132">end</span></span>](#end-datetime) | <span data-ttu-id="78f1e-133">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-133">Member</span></span> |
| [<span data-ttu-id="78f1e-134">from</span><span class="sxs-lookup"><span data-stu-id="78f1e-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="78f1e-135">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-135">Member</span></span> |
| [<span data-ttu-id="78f1e-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="78f1e-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="78f1e-137">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-137">Member</span></span> |
| [<span data-ttu-id="78f1e-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="78f1e-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="78f1e-139">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-139">Member</span></span> |
| [<span data-ttu-id="78f1e-140">itemId</span><span class="sxs-lookup"><span data-stu-id="78f1e-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="78f1e-141">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-141">Member</span></span> |
| [<span data-ttu-id="78f1e-142">itemType</span><span class="sxs-lookup"><span data-stu-id="78f1e-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="78f1e-143">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-143">Member</span></span> |
| [<span data-ttu-id="78f1e-144">location</span><span class="sxs-lookup"><span data-stu-id="78f1e-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="78f1e-145">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-145">Member</span></span> |
| [<span data-ttu-id="78f1e-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="78f1e-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="78f1e-147">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-147">Member</span></span> |
| [<span data-ttu-id="78f1e-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="78f1e-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="78f1e-149">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-149">Member</span></span> |
| [<span data-ttu-id="78f1e-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="78f1e-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="78f1e-151">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-151">Member</span></span> |
| [<span data-ttu-id="78f1e-152">organizer</span><span class="sxs-lookup"><span data-stu-id="78f1e-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="78f1e-153">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-153">Member</span></span> |
| [<span data-ttu-id="78f1e-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="78f1e-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="78f1e-155">Member</span><span class="sxs-lookup"><span data-stu-id="78f1e-155">Member</span></span> |
| [<span data-ttu-id="78f1e-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="78f1e-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="78f1e-157">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-157">Member</span></span> |
| [<span data-ttu-id="78f1e-158">sender</span><span class="sxs-lookup"><span data-stu-id="78f1e-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="78f1e-159">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-159">Member</span></span> |
| [<span data-ttu-id="78f1e-160">seriesid</span><span class="sxs-lookup"><span data-stu-id="78f1e-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="78f1e-161">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-161">Member</span></span> |
| [<span data-ttu-id="78f1e-162">start</span><span class="sxs-lookup"><span data-stu-id="78f1e-162">start</span></span>](#start-datetime) | <span data-ttu-id="78f1e-163">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-163">Member</span></span> |
| [<span data-ttu-id="78f1e-164">subject</span><span class="sxs-lookup"><span data-stu-id="78f1e-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="78f1e-165">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-165">Member</span></span> |
| [<span data-ttu-id="78f1e-166">to</span><span class="sxs-lookup"><span data-stu-id="78f1e-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="78f1e-167">Membro</span><span class="sxs-lookup"><span data-stu-id="78f1e-167">Member</span></span> |
| [<span data-ttu-id="78f1e-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="78f1e-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="78f1e-169">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-169">Method</span></span> |
| [<span data-ttu-id="78f1e-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="78f1e-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="78f1e-171">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-171">Method</span></span> |
| [<span data-ttu-id="78f1e-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="78f1e-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="78f1e-173">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-173">Method</span></span> |
| [<span data-ttu-id="78f1e-174">close</span><span class="sxs-lookup"><span data-stu-id="78f1e-174">close</span></span>](#close) | <span data-ttu-id="78f1e-175">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-175">Method</span></span> |
| [<span data-ttu-id="78f1e-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="78f1e-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="78f1e-177">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-177">Method</span></span> |
| [<span data-ttu-id="78f1e-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="78f1e-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="78f1e-179">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-179">Method</span></span> |
| [<span data-ttu-id="78f1e-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="78f1e-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="78f1e-181">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-181">Method</span></span> |
| [<span data-ttu-id="78f1e-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="78f1e-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="78f1e-183">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-183">Method</span></span> |
| [<span data-ttu-id="78f1e-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="78f1e-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="78f1e-185">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-185">Method</span></span> |
| [<span data-ttu-id="78f1e-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="78f1e-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="78f1e-187">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-187">Method</span></span> |
| [<span data-ttu-id="78f1e-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="78f1e-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="78f1e-189">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-189">Method</span></span> |
| [<span data-ttu-id="78f1e-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="78f1e-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="78f1e-191">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-191">Method</span></span> |
| [<span data-ttu-id="78f1e-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="78f1e-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="78f1e-193">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-193">Method</span></span> |
| [<span data-ttu-id="78f1e-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="78f1e-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="78f1e-195">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-195">Method</span></span> |
| [<span data-ttu-id="78f1e-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="78f1e-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="78f1e-197">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-197">Method</span></span> |
| [<span data-ttu-id="78f1e-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="78f1e-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="78f1e-199">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-199">Method</span></span> |
| [<span data-ttu-id="78f1e-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="78f1e-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="78f1e-201">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-201">Method</span></span> |
| [<span data-ttu-id="78f1e-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="78f1e-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="78f1e-203">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-203">Method</span></span> |
| [<span data-ttu-id="78f1e-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="78f1e-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="78f1e-205">Método</span><span class="sxs-lookup"><span data-stu-id="78f1e-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="78f1e-206">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-206">Example</span></span>

<span data-ttu-id="78f1e-207">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="78f1e-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="78f1e-208">Members</span><span class="sxs-lookup"><span data-stu-id="78f1e-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-17"></a><span data-ttu-id="78f1e-209">anexos: Matriz.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="78f1e-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

<span data-ttu-id="78f1e-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-212">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="78f1e-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="78f1e-213">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="78f1e-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-214">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-214">Type</span></span>

*   <span data-ttu-id="78f1e-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="78f1e-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-216">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-216">Requirements</span></span>

|<span data-ttu-id="78f1e-217">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-217">Requirement</span></span>|<span data-ttu-id="78f1e-218">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-219">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-220">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-220">1.0</span></span>|
|[<span data-ttu-id="78f1e-221">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-222">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-224">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-225">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-225">Example</span></span>

<span data-ttu-id="78f1e-226">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="78f1e-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="78f1e-227">cco :[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-228">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="78f1e-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="78f1e-229">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="78f1e-229">Compose mode only.</span></span>

<span data-ttu-id="78f1e-230">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-230">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="78f1e-231">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="78f1e-231">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="78f1e-232">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-232">Get 500 members maximum.</span></span>
- <span data-ttu-id="78f1e-233">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="78f1e-233">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-234">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-234">Type</span></span>

*   [<span data-ttu-id="78f1e-235">Destinatários</span><span class="sxs-lookup"><span data-stu-id="78f1e-235">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="78f1e-236">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-236">Requirements</span></span>

|<span data-ttu-id="78f1e-237">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-237">Requirement</span></span>|<span data-ttu-id="78f1e-238">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-239">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-240">1.1</span><span class="sxs-lookup"><span data-stu-id="78f1e-240">1.1</span></span>|
|[<span data-ttu-id="78f1e-241">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-242">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-243">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-244">Escrever</span><span class="sxs-lookup"><span data-stu-id="78f1e-244">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-245">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-245">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-17"></a><span data-ttu-id="78f1e-246">corpo: [Corpo](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-247">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="78f1e-247">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-248">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-248">Type</span></span>

*   [<span data-ttu-id="78f1e-249">Body</span><span class="sxs-lookup"><span data-stu-id="78f1e-249">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="78f1e-250">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-250">Requirements</span></span>

|<span data-ttu-id="78f1e-251">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-251">Requirement</span></span>|<span data-ttu-id="78f1e-252">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-253">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-254">1.1</span><span class="sxs-lookup"><span data-stu-id="78f1e-254">1.1</span></span>|
|[<span data-ttu-id="78f1e-255">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-256">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-257">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-258">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-259">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-259">Example</span></span>

<span data-ttu-id="78f1e-260">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="78f1e-260">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="78f1e-261">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-261">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="78f1e-262">cc : Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-263">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="78f1e-263">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="78f1e-264">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="78f1e-264">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="78f1e-265">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="78f1e-265">Read mode</span></span>

<span data-ttu-id="78f1e-266">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="78f1e-266">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="78f1e-267">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-267">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="78f1e-268">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-268">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="78f1e-269">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="78f1e-269">Compose mode</span></span>

<span data-ttu-id="78f1e-270">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="78f1e-270">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="78f1e-271">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-271">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="78f1e-272">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="78f1e-272">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="78f1e-273">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-273">Get 500 members maximum.</span></span>
- <span data-ttu-id="78f1e-274">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="78f1e-274">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="78f1e-275">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-275">Type</span></span>

*   <span data-ttu-id="78f1e-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-277">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-277">Requirements</span></span>

|<span data-ttu-id="78f1e-278">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-278">Requirement</span></span>|<span data-ttu-id="78f1e-279">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-280">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-281">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-281">1.0</span></span>|
|[<span data-ttu-id="78f1e-282">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-283">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-284">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-285">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-285">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="78f1e-286">(anulável) conversationId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-286">(nullable) conversationId: String</span></span>

<span data-ttu-id="78f1e-287">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="78f1e-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="78f1e-p109">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="78f1e-p110">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-292">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-292">Type</span></span>

*   <span data-ttu-id="78f1e-293">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-294">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-294">Requirements</span></span>

|<span data-ttu-id="78f1e-295">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-295">Requirement</span></span>|<span data-ttu-id="78f1e-296">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-297">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-298">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-298">1.0</span></span>|
|[<span data-ttu-id="78f1e-299">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-299">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-300">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-301">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-301">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-302">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-302">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-303">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-303">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="78f1e-304">dateTimeCreated: Data</span><span class="sxs-lookup"><span data-stu-id="78f1e-304">dateTimeCreated: Date</span></span>

<span data-ttu-id="78f1e-p111">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-307">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-307">Type</span></span>

*   <span data-ttu-id="78f1e-308">Data</span><span class="sxs-lookup"><span data-stu-id="78f1e-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-309">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-309">Requirements</span></span>

|<span data-ttu-id="78f1e-310">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-310">Requirement</span></span>|<span data-ttu-id="78f1e-311">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-312">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-313">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-313">1.0</span></span>|
|[<span data-ttu-id="78f1e-314">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-315">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-316">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-317">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-318">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-318">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="78f1e-319">dateTimeModified: Data</span><span class="sxs-lookup"><span data-stu-id="78f1e-319">dateTimeModified: Date</span></span>

<span data-ttu-id="78f1e-p112">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-322">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="78f1e-322">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-323">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-323">Type</span></span>

*   <span data-ttu-id="78f1e-324">Data</span><span class="sxs-lookup"><span data-stu-id="78f1e-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-325">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-325">Requirements</span></span>

|<span data-ttu-id="78f1e-326">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-326">Requirement</span></span>|<span data-ttu-id="78f1e-327">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-328">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-329">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-329">1.0</span></span>|
|[<span data-ttu-id="78f1e-330">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-331">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-332">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-333">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-334">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-334">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="78f1e-335">fim: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-336">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="78f1e-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="78f1e-p113">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="78f1e-339">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="78f1e-339">Read mode</span></span>

<span data-ttu-id="78f1e-340">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-340">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="78f1e-341">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="78f1e-341">Compose mode</span></span>

<span data-ttu-id="78f1e-342">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="78f1e-343">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="78f1e-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="78f1e-344">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-344">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="78f1e-345">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-345">Type</span></span>

*   <span data-ttu-id="78f1e-346">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-347">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-347">Requirements</span></span>

|<span data-ttu-id="78f1e-348">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-348">Requirement</span></span>|<span data-ttu-id="78f1e-349">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-350">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-351">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-351">1.0</span></span>|
|[<span data-ttu-id="78f1e-352">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-353">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-354">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-355">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-355">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17fromjavascriptapioutlookofficefromviewoutlook-js-17"></a><span data-ttu-id="78f1e-356">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[de](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-356">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-357">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="78f1e-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="78f1e-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-360">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-360">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="78f1e-361">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="78f1e-361">Read mode</span></span>

<span data-ttu-id="78f1e-362">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="78f1e-362">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="78f1e-363">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="78f1e-363">Compose mode</span></span>

<span data-ttu-id="78f1e-364">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="78f1e-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="78f1e-365">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-365">Type</span></span>

*   <span data-ttu-id="78f1e-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [de](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-367">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-367">Requirements</span></span>

|<span data-ttu-id="78f1e-368">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="78f1e-369">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-370">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-370">1.0</span></span>|<span data-ttu-id="78f1e-371">1.7</span><span class="sxs-lookup"><span data-stu-id="78f1e-371">1.7</span></span>|
|[<span data-ttu-id="78f1e-372">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-373">ReadItem</span></span>|<span data-ttu-id="78f1e-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="78f1e-375">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-375">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-376">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-376">Read</span></span>|<span data-ttu-id="78f1e-377">Escrever</span><span class="sxs-lookup"><span data-stu-id="78f1e-377">Compose</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="78f1e-378">internetMessageId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-378">internetMessageId: String</span></span>

<span data-ttu-id="78f1e-p115">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-381">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-381">Type</span></span>

*   <span data-ttu-id="78f1e-382">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-383">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-383">Requirements</span></span>

|<span data-ttu-id="78f1e-384">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-384">Requirement</span></span>|<span data-ttu-id="78f1e-385">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-386">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-387">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-387">1.0</span></span>|
|[<span data-ttu-id="78f1e-388">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-389">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-390">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-391">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-392">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-392">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="78f1e-393">itemClass: Cadeia de caracteres </span><span class="sxs-lookup"><span data-stu-id="78f1e-393">itemClass: String</span></span>

<span data-ttu-id="78f1e-p116">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="78f1e-p117">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="78f1e-398">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-398">Type</span></span>|<span data-ttu-id="78f1e-399">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-399">Description</span></span>|<span data-ttu-id="78f1e-400">classe de item</span><span class="sxs-lookup"><span data-stu-id="78f1e-400">item class</span></span>|
|---|---|---|
|<span data-ttu-id="78f1e-401">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="78f1e-401">Appointment items</span></span>|<span data-ttu-id="78f1e-402">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-402">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="78f1e-403">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="78f1e-403">Message items</span></span>|<span data-ttu-id="78f1e-404">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="78f1e-404">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="78f1e-405">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-405">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-406">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-406">Type</span></span>

*   <span data-ttu-id="78f1e-407">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-408">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-408">Requirements</span></span>

|<span data-ttu-id="78f1e-409">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-409">Requirement</span></span>|<span data-ttu-id="78f1e-410">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-411">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-412">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-412">1.0</span></span>|
|[<span data-ttu-id="78f1e-413">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-414">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-415">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-416">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-417">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-417">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="78f1e-418">(anulável) itemId: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-418">(nullable) itemId: String</span></span>

<span data-ttu-id="78f1e-p118">Obtém o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p118">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-421">O identificador retornado pela propriedade `itemId` é o mesmo que o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="78f1e-421">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="78f1e-422">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="78f1e-422">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="78f1e-423">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="78f1e-423">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="78f1e-424">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="78f1e-424">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="78f1e-p120">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p120">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-427">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-427">Type</span></span>

*   <span data-ttu-id="78f1e-428">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-428">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-429">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-429">Requirements</span></span>

|<span data-ttu-id="78f1e-430">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-430">Requirement</span></span>|<span data-ttu-id="78f1e-431">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-431">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-432">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-432">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-433">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-433">1.0</span></span>|
|[<span data-ttu-id="78f1e-434">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-434">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-435">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-435">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-436">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-436">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-437">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-437">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-438">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-438">Example</span></span>

<span data-ttu-id="78f1e-p121">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p121">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-17"></a><span data-ttu-id="78f1e-441">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-442">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="78f1e-442">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="78f1e-443">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="78f1e-443">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-444">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-444">Type</span></span>

*   [<span data-ttu-id="78f1e-445">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="78f1e-445">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="78f1e-446">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-446">Requirements</span></span>

|<span data-ttu-id="78f1e-447">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-447">Requirement</span></span>|<span data-ttu-id="78f1e-448">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-449">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-450">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-450">1.0</span></span>|
|[<span data-ttu-id="78f1e-451">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-451">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-452">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-453">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-453">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-454">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-454">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-455">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-455">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-17"></a><span data-ttu-id="78f1e-456">Local: Cadeia de caracteres[Local](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-457">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="78f1e-457">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="78f1e-458">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="78f1e-458">Read mode</span></span>

<span data-ttu-id="78f1e-459">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="78f1e-459">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="78f1e-460">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="78f1e-460">Compose mode</span></span>

<span data-ttu-id="78f1e-461">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="78f1e-461">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="78f1e-462">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-462">Type</span></span>

*   <span data-ttu-id="78f1e-463">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-464">Requirements</span></span>

|<span data-ttu-id="78f1e-465">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-465">Requirement</span></span>|<span data-ttu-id="78f1e-466">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-467">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-468">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-468">1.0</span></span>|
|[<span data-ttu-id="78f1e-469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-470">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-471">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-472">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-472">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="78f1e-473">normalizedSubject: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-473">normalizedSubject: String</span></span>

<span data-ttu-id="78f1e-p122">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p122">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="78f1e-p123">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="78f1e-p123">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-478">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-478">Type</span></span>

*   <span data-ttu-id="78f1e-479">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-479">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-480">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-480">Requirements</span></span>

|<span data-ttu-id="78f1e-481">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-481">Requirement</span></span>|<span data-ttu-id="78f1e-482">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-482">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-483">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-484">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-484">1.0</span></span>|
|[<span data-ttu-id="78f1e-485">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-485">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-486">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-486">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-487">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-487">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-488">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-488">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-489">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-489">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-17"></a><span data-ttu-id="78f1e-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-491">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="78f1e-491">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-492">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-492">Type</span></span>

*   [<span data-ttu-id="78f1e-493">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="78f1e-493">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="78f1e-494">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-494">Requirements</span></span>

|<span data-ttu-id="78f1e-495">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-495">Requirement</span></span>|<span data-ttu-id="78f1e-496">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-497">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-498">1.3</span><span class="sxs-lookup"><span data-stu-id="78f1e-498">1.3</span></span>|
|[<span data-ttu-id="78f1e-499">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-499">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-500">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-501">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-501">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-502">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-502">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-503">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-503">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="78f1e-504">optionalAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-505">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="78f1e-505">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="78f1e-506">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="78f1e-506">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="78f1e-507">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="78f1e-507">Read mode</span></span>

<span data-ttu-id="78f1e-508">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="78f1e-508">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="78f1e-509">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-509">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="78f1e-510">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-510">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="78f1e-511">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="78f1e-511">Compose mode</span></span>

<span data-ttu-id="78f1e-512">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="78f1e-512">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="78f1e-513">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-513">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="78f1e-514">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="78f1e-514">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="78f1e-515">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-515">Get 500 members maximum.</span></span>
- <span data-ttu-id="78f1e-516">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="78f1e-516">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="78f1e-517">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-517">Type</span></span>

*   <span data-ttu-id="78f1e-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-519">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-519">Requirements</span></span>

|<span data-ttu-id="78f1e-520">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-520">Requirement</span></span>|<span data-ttu-id="78f1e-521">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-522">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-523">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-523">1.0</span></span>|
|[<span data-ttu-id="78f1e-524">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-525">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-526">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-527">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-527">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17organizerjavascriptapioutlookofficeorganizerviewoutlook-js-17"></a><span data-ttu-id="78f1e-528">organizador: [](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="78f1e-528">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-529">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-529">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="78f1e-530">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="78f1e-530">Read mode</span></span>

<span data-ttu-id="78f1e-531">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="78f1e-531">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="78f1e-532">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="78f1e-532">Compose mode</span></span>

<span data-ttu-id="78f1e-533">A `organizer` propriedade retorna um objeto [organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="78f1e-533">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="78f1e-534">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-534">Type</span></span>

*   <span data-ttu-id="78f1e-535">[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizador](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="78f1e-535">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-536">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-536">Requirements</span></span>

|<span data-ttu-id="78f1e-537">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-537">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="78f1e-538">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-539">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-539">1.0</span></span>|<span data-ttu-id="78f1e-540">1.7</span><span class="sxs-lookup"><span data-stu-id="78f1e-540">1.7</span></span>|
|[<span data-ttu-id="78f1e-541">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-542">ReadItem</span></span>|<span data-ttu-id="78f1e-543">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-543">ReadWriteItem</span></span>|
|[<span data-ttu-id="78f1e-544">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-545">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-545">Read</span></span>|<span data-ttu-id="78f1e-546">Escrever</span><span class="sxs-lookup"><span data-stu-id="78f1e-546">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-17"></a><span data-ttu-id="78f1e-547">(anulável) recorrência: [recorrência](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-547">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-548">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="78f1e-548">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="78f1e-549">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="78f1e-549">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="78f1e-550">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="78f1e-550">Read and compose modes for appointment items.</span></span> <span data-ttu-id="78f1e-551">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="78f1e-551">Read mode for meeting request items.</span></span>

<span data-ttu-id="78f1e-552">A `recurrence` propriedade retorna um objeto de [recorrência](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="78f1e-552">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="78f1e-553">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="78f1e-553">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="78f1e-554">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="78f1e-554">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="78f1e-555">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="78f1e-555">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="78f1e-556">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="78f1e-556">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="78f1e-557">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="78f1e-557">Read mode</span></span>

<span data-ttu-id="78f1e-558">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="78f1e-558">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that represents the appointment recurrence.</span></span> <span data-ttu-id="78f1e-559">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="78f1e-559">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="78f1e-560">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="78f1e-560">Compose mode</span></span>

<span data-ttu-id="78f1e-561">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="78f1e-561">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="78f1e-562">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="78f1e-562">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="78f1e-563">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-563">Type</span></span>

* [<span data-ttu-id="78f1e-564">Recorrência</span><span class="sxs-lookup"><span data-stu-id="78f1e-564">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)

|<span data-ttu-id="78f1e-565">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-565">Requirement</span></span>|<span data-ttu-id="78f1e-566">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-567">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-568">1.7</span><span class="sxs-lookup"><span data-stu-id="78f1e-568">1.7</span></span>|
|[<span data-ttu-id="78f1e-569">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-570">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-571">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-572">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-572">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="78f1e-573">requiredAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-574">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="78f1e-574">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="78f1e-575">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="78f1e-575">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="78f1e-576">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="78f1e-576">Read mode</span></span>

<span data-ttu-id="78f1e-577">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="78f1e-577">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="78f1e-578">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-578">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="78f1e-579">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-579">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="78f1e-580">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="78f1e-580">Compose mode</span></span>

<span data-ttu-id="78f1e-581">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="78f1e-581">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="78f1e-582">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-582">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="78f1e-583">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="78f1e-583">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="78f1e-584">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-584">Get 500 members maximum.</span></span>
- <span data-ttu-id="78f1e-585">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="78f1e-585">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="78f1e-586">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-586">Type</span></span>

*   <span data-ttu-id="78f1e-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-588">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-588">Requirements</span></span>

|<span data-ttu-id="78f1e-589">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-589">Requirement</span></span>|<span data-ttu-id="78f1e-590">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-590">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-591">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-591">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-592">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-592">1.0</span></span>|
|[<span data-ttu-id="78f1e-593">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-594">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-595">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-595">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-596">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-596">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17"></a><span data-ttu-id="78f1e-597">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-p134">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p134">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="78f1e-p135">As propriedades [`from`](#from-emailaddressdetailsfrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p135">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-602">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-602">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-603">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-603">Type</span></span>

*   [<span data-ttu-id="78f1e-604">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="78f1e-604">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="78f1e-605">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-605">Requirements</span></span>

|<span data-ttu-id="78f1e-606">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-606">Requirement</span></span>|<span data-ttu-id="78f1e-607">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-608">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-609">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-609">1.0</span></span>|
|[<span data-ttu-id="78f1e-610">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-611">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-612">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-613">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-613">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-614">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-614">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="78f1e-615">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="78f1e-615">(nullable) seriesId: String</span></span>

<span data-ttu-id="78f1e-616">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="78f1e-616">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="78f1e-617">No Outlook na Web e clientes de desktop, o `seriesId` retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="78f1e-617">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="78f1e-618">No entanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="78f1e-618">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-619">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="78f1e-619">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="78f1e-620">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="78f1e-620">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="78f1e-621">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="78f1e-621">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="78f1e-622">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="78f1e-622">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="78f1e-623">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="78f1e-623">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="78f1e-624">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-624">Type</span></span>

* <span data-ttu-id="78f1e-625">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-625">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-626">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-626">Requirements</span></span>

|<span data-ttu-id="78f1e-627">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-627">Requirement</span></span>|<span data-ttu-id="78f1e-628">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-629">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-630">1.7</span><span class="sxs-lookup"><span data-stu-id="78f1e-630">1.7</span></span>|
|[<span data-ttu-id="78f1e-631">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-631">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-632">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-633">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-633">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-634">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-634">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-635">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-635">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="78f1e-636">início: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-637">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="78f1e-637">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="78f1e-p138">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p138">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="78f1e-640">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="78f1e-640">Read mode</span></span>

<span data-ttu-id="78f1e-641">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-641">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="78f1e-642">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="78f1e-642">Compose mode</span></span>

<span data-ttu-id="78f1e-643">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-643">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="78f1e-644">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="78f1e-644">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="78f1e-645">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-645">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="78f1e-646">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-646">Type</span></span>

*   <span data-ttu-id="78f1e-647">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-648">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-648">Requirements</span></span>

|<span data-ttu-id="78f1e-649">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-649">Requirement</span></span>|<span data-ttu-id="78f1e-650">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-651">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-652">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-652">1.0</span></span>|
|[<span data-ttu-id="78f1e-653">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-653">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-654">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-654">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-655">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-655">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-656">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-656">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-17"></a><span data-ttu-id="78f1e-657">Assunto: Cadeia de caracteres|[Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-658">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="78f1e-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="78f1e-659">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="78f1e-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="78f1e-660">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="78f1e-660">Read mode</span></span>

<span data-ttu-id="78f1e-p139">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p139">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="78f1e-663">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="78f1e-663">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="78f1e-664">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="78f1e-664">Compose mode</span></span>

<span data-ttu-id="78f1e-665">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="78f1e-665">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="78f1e-666">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-666">Type</span></span>

*   <span data-ttu-id="78f1e-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-668">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-668">Requirements</span></span>

|<span data-ttu-id="78f1e-669">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-669">Requirement</span></span>|<span data-ttu-id="78f1e-670">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-670">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-671">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-671">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-672">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-672">1.0</span></span>|
|[<span data-ttu-id="78f1e-673">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-673">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-674">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-674">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-675">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-675">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-676">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-676">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="78f1e-677">para: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="78f1e-678">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="78f1e-678">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="78f1e-679">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="78f1e-679">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="78f1e-680">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="78f1e-680">Read mode</span></span>

<span data-ttu-id="78f1e-681">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="78f1e-681">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="78f1e-682">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-682">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="78f1e-683">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-683">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="78f1e-684">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="78f1e-684">Compose mode</span></span>

<span data-ttu-id="78f1e-685">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="78f1e-685">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="78f1e-686">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-686">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="78f1e-687">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="78f1e-687">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="78f1e-688">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="78f1e-688">Get 500 members maximum.</span></span>
- <span data-ttu-id="78f1e-689">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="78f1e-689">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="78f1e-690">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-690">Type</span></span>

*   <span data-ttu-id="78f1e-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-692">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-692">Requirements</span></span>

|<span data-ttu-id="78f1e-693">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-693">Requirement</span></span>|<span data-ttu-id="78f1e-694">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-695">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-696">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-696">1.0</span></span>|
|[<span data-ttu-id="78f1e-697">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-698">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-699">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-700">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-700">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="78f1e-701">Métodos</span><span class="sxs-lookup"><span data-stu-id="78f1e-701">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="78f1e-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="78f1e-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="78f1e-703">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="78f1e-703">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="78f1e-704">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="78f1e-704">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="78f1e-705">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="78f1e-705">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-706">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-706">Parameters</span></span>
|<span data-ttu-id="78f1e-707">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-707">Name</span></span>|<span data-ttu-id="78f1e-708">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-708">Type</span></span>|<span data-ttu-id="78f1e-709">Atributos</span><span class="sxs-lookup"><span data-stu-id="78f1e-709">Attributes</span></span>|<span data-ttu-id="78f1e-710">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-710">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="78f1e-711">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-711">String</span></span>||<span data-ttu-id="78f1e-p143">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p143">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="78f1e-714">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-714">String</span></span>||<span data-ttu-id="78f1e-p144">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p144">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="78f1e-717">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-717">Object</span></span>|<span data-ttu-id="78f1e-718">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-718">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-719">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="78f1e-719">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="78f1e-720">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-720">Object</span></span>|<span data-ttu-id="78f1e-721">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-721">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-722">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-722">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="78f1e-723">Booliano</span><span class="sxs-lookup"><span data-stu-id="78f1e-723">Boolean</span></span>|<span data-ttu-id="78f1e-724">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-724">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-725">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="78f1e-725">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="78f1e-726">function</span><span class="sxs-lookup"><span data-stu-id="78f1e-726">function</span></span>|<span data-ttu-id="78f1e-727">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-727">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-728">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="78f1e-728">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="78f1e-729">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-729">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="78f1e-730">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="78f1e-730">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="78f1e-731">Erros</span><span class="sxs-lookup"><span data-stu-id="78f1e-731">Errors</span></span>

|<span data-ttu-id="78f1e-732">Código de erro</span><span class="sxs-lookup"><span data-stu-id="78f1e-732">Error code</span></span>|<span data-ttu-id="78f1e-733">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-733">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="78f1e-734">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="78f1e-734">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="78f1e-735">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="78f1e-735">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="78f1e-736">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="78f1e-736">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-737">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-737">Requirements</span></span>

|<span data-ttu-id="78f1e-738">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-738">Requirement</span></span>|<span data-ttu-id="78f1e-739">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-740">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-741">1.1</span><span class="sxs-lookup"><span data-stu-id="78f1e-741">1.1</span></span>|
|[<span data-ttu-id="78f1e-742">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-743">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-743">ReadWriteItem</span></span>|
|[<span data-ttu-id="78f1e-744">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-745">Escrever</span><span class="sxs-lookup"><span data-stu-id="78f1e-745">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="78f1e-746">Exemplos</span><span class="sxs-lookup"><span data-stu-id="78f1e-746">Examples</span></span>

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

<span data-ttu-id="78f1e-747">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="78f1e-747">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="78f1e-748">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="78f1e-748">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="78f1e-749">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="78f1e-749">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="78f1e-750">Atualmente, os tipos de eventos `Office.EventType.AppointmentTimeChanged`suportados `Office.EventType.RecipientsChanged`são, e`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="78f1e-750">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-751">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-751">Parameters</span></span>

| <span data-ttu-id="78f1e-752">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-752">Name</span></span> | <span data-ttu-id="78f1e-753">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-753">Type</span></span> | <span data-ttu-id="78f1e-754">Atributos</span><span class="sxs-lookup"><span data-stu-id="78f1e-754">Attributes</span></span> | <span data-ttu-id="78f1e-755">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-755">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="78f1e-756">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="78f1e-756">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="78f1e-757">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="78f1e-757">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="78f1e-758">Função</span><span class="sxs-lookup"><span data-stu-id="78f1e-758">Function</span></span> || <span data-ttu-id="78f1e-p145">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p145">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="78f1e-762">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-762">Object</span></span> | <span data-ttu-id="78f1e-763">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-763">&lt;optional&gt;</span></span> | <span data-ttu-id="78f1e-764">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="78f1e-764">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="78f1e-765">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-765">Object</span></span> | <span data-ttu-id="78f1e-766">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-766">&lt;optional&gt;</span></span> | <span data-ttu-id="78f1e-767">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-767">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="78f1e-768">function</span><span class="sxs-lookup"><span data-stu-id="78f1e-768">function</span></span>| <span data-ttu-id="78f1e-769">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-769">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-770">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="78f1e-770">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-771">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-771">Requirements</span></span>

|<span data-ttu-id="78f1e-772">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-772">Requirement</span></span>| <span data-ttu-id="78f1e-773">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-773">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-774">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-774">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="78f1e-775">1.7</span><span class="sxs-lookup"><span data-stu-id="78f1e-775">1.7</span></span> |
|[<span data-ttu-id="78f1e-776">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-776">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="78f1e-777">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-777">ReadItem</span></span> |
|[<span data-ttu-id="78f1e-778">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-778">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="78f1e-779">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-779">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="78f1e-780">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-780">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="78f1e-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="78f1e-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="78f1e-782">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="78f1e-782">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="78f1e-p146">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p146">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="78f1e-786">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="78f1e-786">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="78f1e-787">Se o Suplemento do Office estiver em execução no Outlook na Web, o método `addItemAttachmentAsync` pode anexar itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="78f1e-787">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-788">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-788">Parameters</span></span>

|<span data-ttu-id="78f1e-789">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-789">Name</span></span>|<span data-ttu-id="78f1e-790">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-790">Type</span></span>|<span data-ttu-id="78f1e-791">Atributos</span><span class="sxs-lookup"><span data-stu-id="78f1e-791">Attributes</span></span>|<span data-ttu-id="78f1e-792">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-792">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="78f1e-793">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-793">String</span></span>||<span data-ttu-id="78f1e-p147">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p147">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="78f1e-796">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-796">String</span></span>||<span data-ttu-id="78f1e-797">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="78f1e-797">The subject of the item to be attached.</span></span> <span data-ttu-id="78f1e-798">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="78f1e-798">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="78f1e-799">Object</span><span class="sxs-lookup"><span data-stu-id="78f1e-799">Object</span></span>|<span data-ttu-id="78f1e-800">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-800">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-801">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="78f1e-801">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="78f1e-802">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-802">Object</span></span>|<span data-ttu-id="78f1e-803">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-803">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-804">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-804">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="78f1e-805">function</span><span class="sxs-lookup"><span data-stu-id="78f1e-805">function</span></span>|<span data-ttu-id="78f1e-806">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-806">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-807">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="78f1e-807">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="78f1e-808">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-808">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="78f1e-809">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="78f1e-809">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="78f1e-810">Erros</span><span class="sxs-lookup"><span data-stu-id="78f1e-810">Errors</span></span>

|<span data-ttu-id="78f1e-811">Código de erro</span><span class="sxs-lookup"><span data-stu-id="78f1e-811">Error code</span></span>|<span data-ttu-id="78f1e-812">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-812">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="78f1e-813">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="78f1e-813">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-814">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-814">Requirements</span></span>

|<span data-ttu-id="78f1e-815">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-815">Requirement</span></span>|<span data-ttu-id="78f1e-816">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-817">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-818">1.1</span><span class="sxs-lookup"><span data-stu-id="78f1e-818">1.1</span></span>|
|[<span data-ttu-id="78f1e-819">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-820">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-820">ReadWriteItem</span></span>|
|[<span data-ttu-id="78f1e-821">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-822">Escrever</span><span class="sxs-lookup"><span data-stu-id="78f1e-822">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-823">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-823">Example</span></span>

<span data-ttu-id="78f1e-824">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-824">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="78f1e-825">close()</span><span class="sxs-lookup"><span data-stu-id="78f1e-825">close()</span></span>

<span data-ttu-id="78f1e-826">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="78f1e-826">Closes the current item that is being composed.</span></span>

<span data-ttu-id="78f1e-p149">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p149">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-829">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="78f1e-829">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="78f1e-830">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="78f1e-830">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-831">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-831">Requirements</span></span>

|<span data-ttu-id="78f1e-832">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-832">Requirement</span></span>|<span data-ttu-id="78f1e-833">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-834">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-835">1.3</span><span class="sxs-lookup"><span data-stu-id="78f1e-835">1.3</span></span>|
|[<span data-ttu-id="78f1e-836">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-837">Restrito</span><span class="sxs-lookup"><span data-stu-id="78f1e-837">Restricted</span></span>|
|[<span data-ttu-id="78f1e-838">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-839">Escrever</span><span class="sxs-lookup"><span data-stu-id="78f1e-839">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="78f1e-840">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="78f1e-840">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="78f1e-841">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="78f1e-841">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-842">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="78f1e-842">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="78f1e-843">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="78f1e-843">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="78f1e-844">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="78f1e-844">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="78f1e-p150">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p150">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-848">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-848">Parameters</span></span>

|<span data-ttu-id="78f1e-849">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-849">Name</span></span>|<span data-ttu-id="78f1e-850">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-850">Type</span></span>|<span data-ttu-id="78f1e-851">Atributos</span><span class="sxs-lookup"><span data-stu-id="78f1e-851">Attributes</span></span>|<span data-ttu-id="78f1e-852">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-852">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="78f1e-853">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="78f1e-853">String &#124; Object</span></span>||<span data-ttu-id="78f1e-p151">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="78f1e-856">**OU**</span><span class="sxs-lookup"><span data-stu-id="78f1e-856">**OR**</span></span><br/><span data-ttu-id="78f1e-p152">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p152">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="78f1e-859">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-859">String</span></span>|<span data-ttu-id="78f1e-860">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-860">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-p153">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="78f1e-863">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-863">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="78f1e-864">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-864">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-865">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="78f1e-865">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="78f1e-866">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-866">String</span></span>||<span data-ttu-id="78f1e-p154">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p154">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="78f1e-869">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-869">String</span></span>||<span data-ttu-id="78f1e-870">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="78f1e-870">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="78f1e-871">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-871">String</span></span>||<span data-ttu-id="78f1e-p155">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p155">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="78f1e-874">Booliano</span><span class="sxs-lookup"><span data-stu-id="78f1e-874">Boolean</span></span>||<span data-ttu-id="78f1e-p156">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p156">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="78f1e-877">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-877">String</span></span>||<span data-ttu-id="78f1e-p157">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p157">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="78f1e-881">function</span><span class="sxs-lookup"><span data-stu-id="78f1e-881">function</span></span>|<span data-ttu-id="78f1e-882">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-882">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-883">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="78f1e-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-884">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-884">Requirements</span></span>

|<span data-ttu-id="78f1e-885">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-885">Requirement</span></span>|<span data-ttu-id="78f1e-886">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-887">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-888">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-888">1.0</span></span>|
|[<span data-ttu-id="78f1e-889">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-890">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-891">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-892">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-892">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="78f1e-893">Exemplos</span><span class="sxs-lookup"><span data-stu-id="78f1e-893">Examples</span></span>

<span data-ttu-id="78f1e-894">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-894">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="78f1e-895">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="78f1e-895">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="78f1e-896">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="78f1e-896">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="78f1e-897">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="78f1e-897">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="78f1e-898">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="78f1e-898">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="78f1e-899">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-899">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="78f1e-900">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="78f1e-900">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="78f1e-901">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="78f1e-901">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-902">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="78f1e-902">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="78f1e-903">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="78f1e-903">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="78f1e-904">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="78f1e-904">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="78f1e-p158">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p158">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-908">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-908">Parameters</span></span>

|<span data-ttu-id="78f1e-909">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-909">Name</span></span>|<span data-ttu-id="78f1e-910">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-910">Type</span></span>|<span data-ttu-id="78f1e-911">Atributos</span><span class="sxs-lookup"><span data-stu-id="78f1e-911">Attributes</span></span>|<span data-ttu-id="78f1e-912">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="78f1e-913">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="78f1e-913">String &#124; Object</span></span>||<span data-ttu-id="78f1e-p159">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p159">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="78f1e-916">**OU**</span><span class="sxs-lookup"><span data-stu-id="78f1e-916">**OR**</span></span><br/><span data-ttu-id="78f1e-p160">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p160">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="78f1e-919">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-919">String</span></span>|<span data-ttu-id="78f1e-920">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-920">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-p161">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p161">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="78f1e-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="78f1e-924">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-924">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-925">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="78f1e-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="78f1e-926">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-926">String</span></span>||<span data-ttu-id="78f1e-p162">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p162">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="78f1e-929">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-929">String</span></span>||<span data-ttu-id="78f1e-930">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="78f1e-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="78f1e-931">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-931">String</span></span>||<span data-ttu-id="78f1e-p163">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p163">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="78f1e-934">Booliano</span><span class="sxs-lookup"><span data-stu-id="78f1e-934">Boolean</span></span>||<span data-ttu-id="78f1e-p164">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p164">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="78f1e-937">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-937">String</span></span>||<span data-ttu-id="78f1e-p165">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p165">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="78f1e-941">function</span><span class="sxs-lookup"><span data-stu-id="78f1e-941">function</span></span>|<span data-ttu-id="78f1e-942">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-942">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-943">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="78f1e-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-944">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-944">Requirements</span></span>

|<span data-ttu-id="78f1e-945">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-945">Requirement</span></span>|<span data-ttu-id="78f1e-946">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-947">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-948">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-948">1.0</span></span>|
|[<span data-ttu-id="78f1e-949">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-950">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-951">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-952">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="78f1e-953">Exemplos</span><span class="sxs-lookup"><span data-stu-id="78f1e-953">Examples</span></span>

<span data-ttu-id="78f1e-954">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-954">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="78f1e-955">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="78f1e-955">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="78f1e-956">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="78f1e-956">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="78f1e-957">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="78f1e-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="78f1e-958">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="78f1e-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="78f1e-959">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="78f1e-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="78f1e-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="78f1e-961">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="78f1e-961">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-962">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="78f1e-962">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-963">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-963">Requirements</span></span>

|<span data-ttu-id="78f1e-964">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-964">Requirement</span></span>|<span data-ttu-id="78f1e-965">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-966">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-966">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-967">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-967">1.0</span></span>|
|[<span data-ttu-id="78f1e-968">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-968">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-969">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-970">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-970">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-971">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-971">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="78f1e-972">Retorna:</span><span class="sxs-lookup"><span data-stu-id="78f1e-972">Returns:</span></span>

<span data-ttu-id="78f1e-973">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-973">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="78f1e-974">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-974">Example</span></span>

<span data-ttu-id="78f1e-975">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="78f1e-975">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="78f1e-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="78f1e-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="78f1e-977">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="78f1e-977">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-978">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="78f1e-978">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-979">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-979">Parameters</span></span>

|<span data-ttu-id="78f1e-980">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-980">Name</span></span>|<span data-ttu-id="78f1e-981">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-981">Type</span></span>|<span data-ttu-id="78f1e-982">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-982">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="78f1e-983">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="78f1e-983">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.7)|<span data-ttu-id="78f1e-984">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="78f1e-984">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-985">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-985">Requirements</span></span>

|<span data-ttu-id="78f1e-986">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-986">Requirement</span></span>|<span data-ttu-id="78f1e-987">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-987">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-988">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-988">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-989">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-989">1.0</span></span>|
|[<span data-ttu-id="78f1e-990">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-990">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-991">Restrito</span><span class="sxs-lookup"><span data-stu-id="78f1e-991">Restricted</span></span>|
|[<span data-ttu-id="78f1e-992">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-992">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-993">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-993">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="78f1e-994">Retorna:</span><span class="sxs-lookup"><span data-stu-id="78f1e-994">Returns:</span></span>

<span data-ttu-id="78f1e-995">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="78f1e-995">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="78f1e-996">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="78f1e-996">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="78f1e-997">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-997">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="78f1e-998">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="78f1e-998">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="78f1e-999">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="78f1e-999">Value of `entityType`</span></span>|<span data-ttu-id="78f1e-1000">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="78f1e-1000">Type of objects in returned array</span></span>|<span data-ttu-id="78f1e-1001">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="78f1e-1001">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="78f1e-1002">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-1002">String</span></span>|<span data-ttu-id="78f1e-1003">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="78f1e-1003">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="78f1e-1004">Contato</span><span class="sxs-lookup"><span data-stu-id="78f1e-1004">Contact</span></span>|<span data-ttu-id="78f1e-1005">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="78f1e-1005">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="78f1e-1006">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-1006">String</span></span>|<span data-ttu-id="78f1e-1007">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="78f1e-1007">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="78f1e-1008">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="78f1e-1008">MeetingSuggestion</span></span>|<span data-ttu-id="78f1e-1009">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="78f1e-1009">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="78f1e-1010">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="78f1e-1010">PhoneNumber</span></span>|<span data-ttu-id="78f1e-1011">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="78f1e-1011">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="78f1e-1012">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="78f1e-1012">TaskSuggestion</span></span>|<span data-ttu-id="78f1e-1013">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="78f1e-1013">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="78f1e-1014">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-1014">String</span></span>|<span data-ttu-id="78f1e-1015">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="78f1e-1015">**Restricted**</span></span>|

<span data-ttu-id="78f1e-1016">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="78f1e-1016">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

##### <a name="example"></a><span data-ttu-id="78f1e-1017">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1017">Example</span></span>

<span data-ttu-id="78f1e-1018">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1018">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="78f1e-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="78f1e-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="78f1e-1020">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1020">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-1021">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1021">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="78f1e-1022">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1022">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-1023">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-1023">Parameters</span></span>

|<span data-ttu-id="78f1e-1024">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-1024">Name</span></span>|<span data-ttu-id="78f1e-1025">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1025">Type</span></span>|<span data-ttu-id="78f1e-1026">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-1026">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="78f1e-1027">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78f1e-1027">String</span></span>|<span data-ttu-id="78f1e-1028">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1028">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-1029">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1029">Requirements</span></span>

|<span data-ttu-id="78f1e-1030">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-1030">Requirement</span></span>|<span data-ttu-id="78f1e-1031">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-1031">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-1032">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-1032">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-1033">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-1033">1.0</span></span>|
|[<span data-ttu-id="78f1e-1034">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1034">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-1035">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-1035">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-1036">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-1036">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-1037">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-1037">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="78f1e-1038">Retorna:</span><span class="sxs-lookup"><span data-stu-id="78f1e-1038">Returns:</span></span>

<span data-ttu-id="78f1e-p167">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p167">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="78f1e-1041">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="78f1e-1041">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="78f1e-1042">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="78f1e-1042">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="78f1e-1043">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1043">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-1044">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1044">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="78f1e-p168">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="78f1e-1048">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="78f1e-1048">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="78f1e-1049">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1049">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="78f1e-p169">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-1053">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1053">Requirements</span></span>

|<span data-ttu-id="78f1e-1054">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-1054">Requirement</span></span>|<span data-ttu-id="78f1e-1055">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-1056">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-1057">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-1057">1.0</span></span>|
|[<span data-ttu-id="78f1e-1058">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1058">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-1059">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-1059">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-1060">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-1060">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-1061">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-1061">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="78f1e-1062">Retorna:</span><span class="sxs-lookup"><span data-stu-id="78f1e-1062">Returns:</span></span>

<span data-ttu-id="78f1e-p170">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="78f1e-1065">Tipo: Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-1065">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="78f1e-1066">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1066">Example</span></span>

<span data-ttu-id="78f1e-1067">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1067">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="78f1e-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="78f1e-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="78f1e-1069">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1069">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-1070">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1070">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="78f1e-1071">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1071">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="78f1e-p171">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-1074">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-1074">Parameters</span></span>

|<span data-ttu-id="78f1e-1075">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-1075">Name</span></span>|<span data-ttu-id="78f1e-1076">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1076">Type</span></span>|<span data-ttu-id="78f1e-1077">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-1077">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="78f1e-1078">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-1078">String</span></span>|<span data-ttu-id="78f1e-1079">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1079">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-1080">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1080">Requirements</span></span>

|<span data-ttu-id="78f1e-1081">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-1081">Requirement</span></span>|<span data-ttu-id="78f1e-1082">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-1083">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-1083">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-1084">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-1084">1.0</span></span>|
|[<span data-ttu-id="78f1e-1085">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1085">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-1086">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-1087">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-1087">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-1088">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-1088">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="78f1e-1089">Retorna:</span><span class="sxs-lookup"><span data-stu-id="78f1e-1089">Returns:</span></span>

<span data-ttu-id="78f1e-1090">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1090">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="78f1e-1091">Tipo: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="78f1e-1091">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="78f1e-1092">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1092">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="78f1e-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="78f1e-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="78f1e-1094">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1094">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="78f1e-1095">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retornará uma cadeia de caracteres vazia para os dados selecionados.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1095">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.</span></span> <span data-ttu-id="78f1e-1096">Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1096">If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-1097">No Outlook na Web, o método retorna a cadeia de caracteres “null” se nenhum texto for selecionado, mas o cursor estiver no corpo.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1097">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="78f1e-1098">Para verificar essa situação, confira o exemplo mais adiante nesta seção.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1098">To check for this situation, see the example later in this section.</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-1099">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-1099">Parameters</span></span>

|<span data-ttu-id="78f1e-1100">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-1100">Name</span></span>|<span data-ttu-id="78f1e-1101">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1101">Type</span></span>|<span data-ttu-id="78f1e-1102">Atributos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1102">Attributes</span></span>|<span data-ttu-id="78f1e-1103">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-1103">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="78f1e-1104">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="78f1e-1104">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="78f1e-p174">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p174">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="78f1e-1108">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-1108">Object</span></span>|<span data-ttu-id="78f1e-1109">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1109">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-1110">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1110">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="78f1e-1111">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-1111">Object</span></span>|<span data-ttu-id="78f1e-1112">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1112">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-1113">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1113">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="78f1e-1114">function</span><span class="sxs-lookup"><span data-stu-id="78f1e-1114">function</span></span>||<span data-ttu-id="78f1e-1115">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="78f1e-1115">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="78f1e-1116">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1116">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="78f1e-1117">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1117">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-1118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1118">Requirements</span></span>

|<span data-ttu-id="78f1e-1119">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-1119">Requirement</span></span>|<span data-ttu-id="78f1e-1120">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-1120">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-1121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-1121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-1122">1.2</span><span class="sxs-lookup"><span data-stu-id="78f1e-1122">1.2</span></span>|
|[<span data-ttu-id="78f1e-1123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-1124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-1124">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-1125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-1125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-1126">Escrever</span><span class="sxs-lookup"><span data-stu-id="78f1e-1126">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="78f1e-1127">Retorna:</span><span class="sxs-lookup"><span data-stu-id="78f1e-1127">Returns:</span></span>

<span data-ttu-id="78f1e-1128">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1128">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="78f1e-1129">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="78f1e-1129">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="78f1e-1130">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1130">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="78f1e-1131">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="78f1e-1131">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="78f1e-1132">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1132">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="78f1e-1133">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="78f1e-1133">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-1134">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1134">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-1135">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1135">Requirements</span></span>

|<span data-ttu-id="78f1e-1136">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-1136">Requirement</span></span>|<span data-ttu-id="78f1e-1137">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-1138">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-1139">1.6</span><span class="sxs-lookup"><span data-stu-id="78f1e-1139">1.6</span></span>|
|[<span data-ttu-id="78f1e-1140">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-1141">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-1141">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-1142">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-1143">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="78f1e-1144">Retorna:</span><span class="sxs-lookup"><span data-stu-id="78f1e-1144">Returns:</span></span>

<span data-ttu-id="78f1e-1145">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="78f1e-1145">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="78f1e-1146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1146">Example</span></span>

<span data-ttu-id="78f1e-1147">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1147">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="78f1e-1148">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="78f1e-1148">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="78f1e-p177">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="78f1e-p177">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-1151">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1151">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="78f1e-p178">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p178">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="78f1e-1155">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="78f1e-1155">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="78f1e-1156">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1156">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="78f1e-p179">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="78f1e-1160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1160">Requirements</span></span>

|<span data-ttu-id="78f1e-1161">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-1161">Requirement</span></span>|<span data-ttu-id="78f1e-1162">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-1162">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-1163">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-1163">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-1164">1.6</span><span class="sxs-lookup"><span data-stu-id="78f1e-1164">1.6</span></span>|
|[<span data-ttu-id="78f1e-1165">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1165">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-1166">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-1166">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-1167">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-1167">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-1168">Read</span><span class="sxs-lookup"><span data-stu-id="78f1e-1168">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="78f1e-1169">Retorna:</span><span class="sxs-lookup"><span data-stu-id="78f1e-1169">Returns:</span></span>

<span data-ttu-id="78f1e-p180">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="78f1e-1172">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1172">Example</span></span>

<span data-ttu-id="78f1e-1173">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1173">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="78f1e-1174">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="78f1e-1174">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="78f1e-1175">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1175">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="78f1e-p181">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p181">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-1179">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-1179">Parameters</span></span>

|<span data-ttu-id="78f1e-1180">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-1180">Name</span></span>|<span data-ttu-id="78f1e-1181">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1181">Type</span></span>|<span data-ttu-id="78f1e-1182">Atributos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1182">Attributes</span></span>|<span data-ttu-id="78f1e-1183">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-1183">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="78f1e-1184">function</span><span class="sxs-lookup"><span data-stu-id="78f1e-1184">function</span></span>||<span data-ttu-id="78f1e-1185">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="78f1e-1185">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="78f1e-1186">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1186">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="78f1e-1187">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1187">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="78f1e-1188">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-1188">Object</span></span>|<span data-ttu-id="78f1e-1189">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1189">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-1190">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1190">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="78f1e-1191">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1191">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-1192">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1192">Requirements</span></span>

|<span data-ttu-id="78f1e-1193">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-1193">Requirement</span></span>|<span data-ttu-id="78f1e-1194">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-1194">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-1195">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-1195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-1196">1.0</span><span class="sxs-lookup"><span data-stu-id="78f1e-1196">1.0</span></span>|
|[<span data-ttu-id="78f1e-1197">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1197">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-1198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-1198">ReadItem</span></span>|
|[<span data-ttu-id="78f1e-1199">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-1199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-1200">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-1200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-1201">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1201">Example</span></span>

<span data-ttu-id="78f1e-p184">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p184">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="78f1e-1205">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="78f1e-1205">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="78f1e-1206">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1206">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="78f1e-1207">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1207">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="78f1e-1208">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1208">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="78f1e-1209">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1209">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="78f1e-1210">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1210">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-1211">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-1211">Parameters</span></span>

|<span data-ttu-id="78f1e-1212">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-1212">Name</span></span>|<span data-ttu-id="78f1e-1213">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1213">Type</span></span>|<span data-ttu-id="78f1e-1214">Atributos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1214">Attributes</span></span>|<span data-ttu-id="78f1e-1215">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-1215">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="78f1e-1216">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-1216">String</span></span>||<span data-ttu-id="78f1e-1217">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1217">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="78f1e-1218">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-1218">Object</span></span>|<span data-ttu-id="78f1e-1219">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1219">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-1220">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1220">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="78f1e-1221">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-1221">Object</span></span>|<span data-ttu-id="78f1e-1222">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1222">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-1223">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1223">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="78f1e-1224">function</span><span class="sxs-lookup"><span data-stu-id="78f1e-1224">function</span></span>|<span data-ttu-id="78f1e-1225">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1225">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-1226">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="78f1e-1226">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="78f1e-1227">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1227">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="78f1e-1228">Erros</span><span class="sxs-lookup"><span data-stu-id="78f1e-1228">Errors</span></span>

|<span data-ttu-id="78f1e-1229">Código de erro</span><span class="sxs-lookup"><span data-stu-id="78f1e-1229">Error code</span></span>|<span data-ttu-id="78f1e-1230">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-1230">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="78f1e-1231">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1231">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-1232">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1232">Requirements</span></span>

|<span data-ttu-id="78f1e-1233">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-1233">Requirement</span></span>|<span data-ttu-id="78f1e-1234">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-1234">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-1235">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-1235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-1236">1.1</span><span class="sxs-lookup"><span data-stu-id="78f1e-1236">1.1</span></span>|
|[<span data-ttu-id="78f1e-1237">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-1238">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-1238">ReadWriteItem</span></span>|
|[<span data-ttu-id="78f1e-1239">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-1239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-1240">Escrever</span><span class="sxs-lookup"><span data-stu-id="78f1e-1240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-1241">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1241">Example</span></span>

<span data-ttu-id="78f1e-1242">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1242">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="78f1e-1243">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="78f1e-1243">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="78f1e-1244">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1244">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="78f1e-1245">Atualmente, os tipos de eventos `Office.EventType.AppointmentTimeChanged`suportados `Office.EventType.RecipientsChanged`são, e`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="78f1e-1245">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-1246">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-1246">Parameters</span></span>

| <span data-ttu-id="78f1e-1247">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-1247">Name</span></span> | <span data-ttu-id="78f1e-1248">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1248">Type</span></span> | <span data-ttu-id="78f1e-1249">Atributos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1249">Attributes</span></span> | <span data-ttu-id="78f1e-1250">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-1250">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="78f1e-1251">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="78f1e-1251">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="78f1e-1252">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1252">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="78f1e-1253">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-1253">Object</span></span> | <span data-ttu-id="78f1e-1254">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1254">&lt;optional&gt;</span></span> | <span data-ttu-id="78f1e-1255">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1255">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="78f1e-1256">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-1256">Object</span></span> | <span data-ttu-id="78f1e-1257">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1257">&lt;optional&gt;</span></span> | <span data-ttu-id="78f1e-1258">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1258">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="78f1e-1259">function</span><span class="sxs-lookup"><span data-stu-id="78f1e-1259">function</span></span>| <span data-ttu-id="78f1e-1260">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1260">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-1261">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="78f1e-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-1262">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1262">Requirements</span></span>

|<span data-ttu-id="78f1e-1263">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-1263">Requirement</span></span>| <span data-ttu-id="78f1e-1264">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-1264">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-1265">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-1265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="78f1e-1266">1.7</span><span class="sxs-lookup"><span data-stu-id="78f1e-1266">1.7</span></span> |
|[<span data-ttu-id="78f1e-1267">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1267">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="78f1e-1268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-1268">ReadItem</span></span> |
|[<span data-ttu-id="78f1e-1269">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="78f1e-1269">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="78f1e-1270">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78f1e-1270">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="78f1e-1271">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1271">Example</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="78f1e-1272">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="78f1e-1272">saveAsync([options], callback)</span></span>

<span data-ttu-id="78f1e-1273">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1273">Asynchronously saves an item.</span></span>

<span data-ttu-id="78f1e-1274">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1274">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="78f1e-1275">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1275">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="78f1e-1276">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1276">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-1277">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1277">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="78f1e-1278">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1278">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="78f1e-p188">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p188">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="78f1e-1282">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="78f1e-1282">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="78f1e-1283">O Outlook no Mac não dá suporte ao salvamento de reuniões.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1283">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="78f1e-1284">O método `saveAsync` falha quando chamado a partir de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1284">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="78f1e-1285">Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1285">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="78f1e-1286">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1286">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-1287">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-1287">Parameters</span></span>

|<span data-ttu-id="78f1e-1288">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-1288">Name</span></span>|<span data-ttu-id="78f1e-1289">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1289">Type</span></span>|<span data-ttu-id="78f1e-1290">Atributos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1290">Attributes</span></span>|<span data-ttu-id="78f1e-1291">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-1291">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="78f1e-1292">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-1292">Object</span></span>|<span data-ttu-id="78f1e-1293">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1293">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-1294">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1294">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="78f1e-1295">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-1295">Object</span></span>|<span data-ttu-id="78f1e-1296">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-1297">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1297">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="78f1e-1298">function</span><span class="sxs-lookup"><span data-stu-id="78f1e-1298">function</span></span>||<span data-ttu-id="78f1e-1299">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="78f1e-1299">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="78f1e-1300">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1300">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-1301">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1301">Requirements</span></span>

|<span data-ttu-id="78f1e-1302">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-1302">Requirement</span></span>|<span data-ttu-id="78f1e-1303">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-1303">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-1304">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-1304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-1305">1.3</span><span class="sxs-lookup"><span data-stu-id="78f1e-1305">1.3</span></span>|
|[<span data-ttu-id="78f1e-1306">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1306">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-1307">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-1307">ReadWriteItem</span></span>|
|[<span data-ttu-id="78f1e-1308">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-1308">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-1309">Escrever</span><span class="sxs-lookup"><span data-stu-id="78f1e-1309">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="78f1e-1310">Exemplos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1310">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="78f1e-p190">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p190">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="78f1e-1313">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="78f1e-1313">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="78f1e-1314">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1314">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="78f1e-p191">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p191">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="78f1e-1318">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="78f1e-1318">Parameters</span></span>

|<span data-ttu-id="78f1e-1319">Nome</span><span class="sxs-lookup"><span data-stu-id="78f1e-1319">Name</span></span>|<span data-ttu-id="78f1e-1320">Tipo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1320">Type</span></span>|<span data-ttu-id="78f1e-1321">Atributos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1321">Attributes</span></span>|<span data-ttu-id="78f1e-1322">Descrição</span><span class="sxs-lookup"><span data-stu-id="78f1e-1322">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="78f1e-1323">String</span><span class="sxs-lookup"><span data-stu-id="78f1e-1323">String</span></span>||<span data-ttu-id="78f1e-p192">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-p192">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="78f1e-1327">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-1327">Object</span></span>|<span data-ttu-id="78f1e-1328">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1328">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-1329">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1329">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="78f1e-1330">Objeto</span><span class="sxs-lookup"><span data-stu-id="78f1e-1330">Object</span></span>|<span data-ttu-id="78f1e-1331">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1331">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-1332">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1332">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="78f1e-1333">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="78f1e-1333">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="78f1e-1334">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="78f1e-1334">&lt;optional&gt;</span></span>|<span data-ttu-id="78f1e-1335">Se `text`, o estilo atual é aplicado nos clientes do Outlook na Web e do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1335">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="78f1e-1336">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1336">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="78f1e-1337">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook na Web e o estilo padrão é aplicado nos clientes do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1337">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="78f1e-1338">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1338">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="78f1e-1339">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="78f1e-1339">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="78f1e-1340">function</span><span class="sxs-lookup"><span data-stu-id="78f1e-1340">function</span></span>||<span data-ttu-id="78f1e-1341">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="78f1e-1341">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78f1e-1342">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78f1e-1342">Requirements</span></span>

|<span data-ttu-id="78f1e-1343">Requisito</span><span class="sxs-lookup"><span data-stu-id="78f1e-1343">Requirement</span></span>|<span data-ttu-id="78f1e-1344">Valor</span><span class="sxs-lookup"><span data-stu-id="78f1e-1344">Value</span></span>|
|---|---|
|[<span data-ttu-id="78f1e-1345">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78f1e-1345">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="78f1e-1346">1.2</span><span class="sxs-lookup"><span data-stu-id="78f1e-1346">1.2</span></span>|
|[<span data-ttu-id="78f1e-1347">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1347">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="78f1e-1348">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="78f1e-1348">ReadWriteItem</span></span>|
|[<span data-ttu-id="78f1e-1349">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78f1e-1349">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="78f1e-1350">Escrever</span><span class="sxs-lookup"><span data-stu-id="78f1e-1350">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="78f1e-1351">Exemplo</span><span class="sxs-lookup"><span data-stu-id="78f1e-1351">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
