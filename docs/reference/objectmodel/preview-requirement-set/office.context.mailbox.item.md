---
title: Office. Context. Mailbox. Item-visualização do conjunto de requisitos
description: ''
ms.date: 09/23/2019
localization_priority: Normal
ms.openlocfilehash: 4a209ebde75a2857f4caa6d246c83adbd2cf7c10
ms.sourcegitcommit: 3c84fe6302341668c3f9f6dd64e636a97d03023c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/26/2019
ms.locfileid: "37167372"
---
# <a name="item"></a><span data-ttu-id="57544-102">item</span><span class="sxs-lookup"><span data-stu-id="57544-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="57544-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="57544-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="57544-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="57544-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-106">Requirements</span></span>

|<span data-ttu-id="57544-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-107">Requirement</span></span>|<span data-ttu-id="57544-108">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-110">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-110">1.0</span></span>|
|[<span data-ttu-id="57544-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="57544-112">Restricted</span></span>|
|[<span data-ttu-id="57544-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="57544-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="57544-115">Members and methods</span></span>

| <span data-ttu-id="57544-116">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-116">Member</span></span> | <span data-ttu-id="57544-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="57544-118">attachments</span><span class="sxs-lookup"><span data-stu-id="57544-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="57544-119">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-119">Member</span></span> |
| [<span data-ttu-id="57544-120">bcc</span><span class="sxs-lookup"><span data-stu-id="57544-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="57544-121">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-121">Member</span></span> |
| [<span data-ttu-id="57544-122">body</span><span class="sxs-lookup"><span data-stu-id="57544-122">body</span></span>](#body-body) | <span data-ttu-id="57544-123">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-123">Member</span></span> |
| [<span data-ttu-id="57544-124">Categorias</span><span class="sxs-lookup"><span data-stu-id="57544-124">categories</span></span>](#categories-categories) | <span data-ttu-id="57544-125">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-125">Member</span></span> |
| [<span data-ttu-id="57544-126">cc</span><span class="sxs-lookup"><span data-stu-id="57544-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="57544-127">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-127">Member</span></span> |
| [<span data-ttu-id="57544-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="57544-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="57544-129">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-129">Member</span></span> |
| [<span data-ttu-id="57544-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="57544-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="57544-131">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-131">Member</span></span> |
| [<span data-ttu-id="57544-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="57544-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="57544-133">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-133">Member</span></span> |
| [<span data-ttu-id="57544-134">end</span><span class="sxs-lookup"><span data-stu-id="57544-134">end</span></span>](#end-datetime) | <span data-ttu-id="57544-135">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-135">Member</span></span> |
| [<span data-ttu-id="57544-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="57544-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="57544-137">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-137">Member</span></span> |
| [<span data-ttu-id="57544-138">from</span><span class="sxs-lookup"><span data-stu-id="57544-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="57544-139">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-139">Member</span></span> |
| [<span data-ttu-id="57544-140">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="57544-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="57544-141">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-141">Member</span></span> |
| [<span data-ttu-id="57544-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="57544-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="57544-143">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-143">Member</span></span> |
| [<span data-ttu-id="57544-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="57544-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="57544-145">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-145">Member</span></span> |
| [<span data-ttu-id="57544-146">itemId</span><span class="sxs-lookup"><span data-stu-id="57544-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="57544-147">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-147">Member</span></span> |
| [<span data-ttu-id="57544-148">itemType</span><span class="sxs-lookup"><span data-stu-id="57544-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="57544-149">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-149">Member</span></span> |
| [<span data-ttu-id="57544-150">location</span><span class="sxs-lookup"><span data-stu-id="57544-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="57544-151">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-151">Member</span></span> |
| [<span data-ttu-id="57544-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="57544-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="57544-153">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-153">Member</span></span> |
| [<span data-ttu-id="57544-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="57544-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="57544-155">Member</span><span class="sxs-lookup"><span data-stu-id="57544-155">Member</span></span> |
| [<span data-ttu-id="57544-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="57544-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="57544-157">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-157">Member</span></span> |
| [<span data-ttu-id="57544-158">organizer</span><span class="sxs-lookup"><span data-stu-id="57544-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="57544-159">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-159">Member</span></span> |
| [<span data-ttu-id="57544-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="57544-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="57544-161">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-161">Member</span></span> |
| [<span data-ttu-id="57544-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="57544-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="57544-163">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-163">Member</span></span> |
| [<span data-ttu-id="57544-164">sender</span><span class="sxs-lookup"><span data-stu-id="57544-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="57544-165">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-165">Member</span></span> |
| [<span data-ttu-id="57544-166">seriesid</span><span class="sxs-lookup"><span data-stu-id="57544-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="57544-167">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-167">Member</span></span> |
| [<span data-ttu-id="57544-168">start</span><span class="sxs-lookup"><span data-stu-id="57544-168">start</span></span>](#start-datetime) | <span data-ttu-id="57544-169">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-169">Member</span></span> |
| [<span data-ttu-id="57544-170">subject</span><span class="sxs-lookup"><span data-stu-id="57544-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="57544-171">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-171">Member</span></span> |
| [<span data-ttu-id="57544-172">to</span><span class="sxs-lookup"><span data-stu-id="57544-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="57544-173">Membro</span><span class="sxs-lookup"><span data-stu-id="57544-173">Member</span></span> |
| [<span data-ttu-id="57544-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="57544-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="57544-175">Método</span><span class="sxs-lookup"><span data-stu-id="57544-175">Method</span></span> |
| [<span data-ttu-id="57544-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="57544-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="57544-177">Método</span><span class="sxs-lookup"><span data-stu-id="57544-177">Method</span></span> |
| [<span data-ttu-id="57544-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="57544-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="57544-179">Método</span><span class="sxs-lookup"><span data-stu-id="57544-179">Method</span></span> |
| [<span data-ttu-id="57544-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="57544-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="57544-181">Método</span><span class="sxs-lookup"><span data-stu-id="57544-181">Method</span></span> |
| [<span data-ttu-id="57544-182">close</span><span class="sxs-lookup"><span data-stu-id="57544-182">close</span></span>](#close) | <span data-ttu-id="57544-183">Método</span><span class="sxs-lookup"><span data-stu-id="57544-183">Method</span></span> |
| [<span data-ttu-id="57544-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="57544-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="57544-185">Método</span><span class="sxs-lookup"><span data-stu-id="57544-185">Method</span></span> |
| [<span data-ttu-id="57544-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="57544-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="57544-187">Método</span><span class="sxs-lookup"><span data-stu-id="57544-187">Method</span></span> |
| [<span data-ttu-id="57544-188">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="57544-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="57544-189">Método</span><span class="sxs-lookup"><span data-stu-id="57544-189">Method</span></span> |
| [<span data-ttu-id="57544-190">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="57544-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="57544-191">Método</span><span class="sxs-lookup"><span data-stu-id="57544-191">Method</span></span> |
| [<span data-ttu-id="57544-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="57544-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="57544-193">Método</span><span class="sxs-lookup"><span data-stu-id="57544-193">Method</span></span> |
| [<span data-ttu-id="57544-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="57544-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="57544-195">Método</span><span class="sxs-lookup"><span data-stu-id="57544-195">Method</span></span> |
| [<span data-ttu-id="57544-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="57544-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="57544-197">Método</span><span class="sxs-lookup"><span data-stu-id="57544-197">Method</span></span> |
| [<span data-ttu-id="57544-198">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="57544-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="57544-199">Método</span><span class="sxs-lookup"><span data-stu-id="57544-199">Method</span></span> |
| [<span data-ttu-id="57544-200">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="57544-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="57544-201">Método</span><span class="sxs-lookup"><span data-stu-id="57544-201">Method</span></span> |
| [<span data-ttu-id="57544-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="57544-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="57544-203">Método</span><span class="sxs-lookup"><span data-stu-id="57544-203">Method</span></span> |
| [<span data-ttu-id="57544-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="57544-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="57544-205">Método</span><span class="sxs-lookup"><span data-stu-id="57544-205">Method</span></span> |
| [<span data-ttu-id="57544-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="57544-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="57544-207">Método</span><span class="sxs-lookup"><span data-stu-id="57544-207">Method</span></span> |
| [<span data-ttu-id="57544-208">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="57544-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="57544-209">Método</span><span class="sxs-lookup"><span data-stu-id="57544-209">Method</span></span> |
| [<span data-ttu-id="57544-210">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="57544-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="57544-211">Método</span><span class="sxs-lookup"><span data-stu-id="57544-211">Method</span></span> |
| [<span data-ttu-id="57544-212">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="57544-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="57544-213">Método</span><span class="sxs-lookup"><span data-stu-id="57544-213">Method</span></span> |
| [<span data-ttu-id="57544-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="57544-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="57544-215">Método</span><span class="sxs-lookup"><span data-stu-id="57544-215">Method</span></span> |
| [<span data-ttu-id="57544-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="57544-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="57544-217">Método</span><span class="sxs-lookup"><span data-stu-id="57544-217">Method</span></span> |
| [<span data-ttu-id="57544-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="57544-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="57544-219">Método</span><span class="sxs-lookup"><span data-stu-id="57544-219">Method</span></span> |
| [<span data-ttu-id="57544-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="57544-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="57544-221">Método</span><span class="sxs-lookup"><span data-stu-id="57544-221">Method</span></span> |
| [<span data-ttu-id="57544-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="57544-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="57544-223">Método</span><span class="sxs-lookup"><span data-stu-id="57544-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="57544-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-224">Example</span></span>

<span data-ttu-id="57544-225">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="57544-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="57544-226">Membros</span><span class="sxs-lookup"><span data-stu-id="57544-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="57544-227">anexos: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="57544-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="57544-228">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="57544-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="57544-229">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="57544-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-230">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="57544-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="57544-231">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="57544-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="57544-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-232">Type</span></span>

*   <span data-ttu-id="57544-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="57544-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-234">Requirements</span></span>

|<span data-ttu-id="57544-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-235">Requirement</span></span>|<span data-ttu-id="57544-236">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-238">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-238">1.0</span></span>|
|[<span data-ttu-id="57544-239">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-240">ReadItem</span></span>|
|[<span data-ttu-id="57544-241">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-242">Read</span><span class="sxs-lookup"><span data-stu-id="57544-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-243">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-243">Example</span></span>

<span data-ttu-id="57544-244">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="57544-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="57544-245">CCO: [destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="57544-245">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="57544-246">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="57544-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="57544-247">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="57544-247">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-248">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-248">Type</span></span>

*   [<span data-ttu-id="57544-249">Destinatários</span><span class="sxs-lookup"><span data-stu-id="57544-249">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="57544-250">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-250">Requirements</span></span>

|<span data-ttu-id="57544-251">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-251">Requirement</span></span>|<span data-ttu-id="57544-252">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-253">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-254">1.1</span><span class="sxs-lookup"><span data-stu-id="57544-254">1.1</span></span>|
|[<span data-ttu-id="57544-255">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-256">ReadItem</span></span>|
|[<span data-ttu-id="57544-257">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-258">Escrever</span><span class="sxs-lookup"><span data-stu-id="57544-258">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-259">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-259">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="57544-260">corpo: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="57544-260">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="57544-261">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="57544-261">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-262">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-262">Type</span></span>

*   [<span data-ttu-id="57544-263">Body</span><span class="sxs-lookup"><span data-stu-id="57544-263">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="57544-264">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-264">Requirements</span></span>

|<span data-ttu-id="57544-265">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-265">Requirement</span></span>|<span data-ttu-id="57544-266">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-267">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-268">1.1</span><span class="sxs-lookup"><span data-stu-id="57544-268">1.1</span></span>|
|[<span data-ttu-id="57544-269">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-270">ReadItem</span></span>|
|[<span data-ttu-id="57544-271">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-272">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-273">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-273">Example</span></span>

<span data-ttu-id="57544-274">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="57544-274">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="57544-275">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-275">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="57544-276">Categorias: [categorias](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="57544-276">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="57544-277">Obtém um objeto que fornece métodos para gerenciar as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="57544-277">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-278">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="57544-278">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-279">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-279">Type</span></span>

*   [<span data-ttu-id="57544-280">Categories</span><span class="sxs-lookup"><span data-stu-id="57544-280">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="57544-281">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-281">Requirements</span></span>

|<span data-ttu-id="57544-282">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-282">Requirement</span></span>|<span data-ttu-id="57544-283">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-284">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-285">Visualização</span><span class="sxs-lookup"><span data-stu-id="57544-285">Preview</span></span>|
|[<span data-ttu-id="57544-286">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-287">ReadItem</span></span>|
|[<span data-ttu-id="57544-288">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-289">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-289">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-290">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-290">Example</span></span>

<span data-ttu-id="57544-291">Este exemplo obtém as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="57544-291">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="57544-292">[destinatários](/javascript/api/outlook/office.recipients) [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|CC: Array. <</span><span class="sxs-lookup"><span data-stu-id="57544-292">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="57544-293">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="57544-293">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="57544-294">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="57544-294">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57544-295">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="57544-295">Read mode</span></span>

<span data-ttu-id="57544-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="57544-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="57544-298">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="57544-298">Compose mode</span></span>

<span data-ttu-id="57544-299">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="57544-299">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="57544-300">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-300">Type</span></span>

*   <span data-ttu-id="57544-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="57544-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-302">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-302">Requirements</span></span>

|<span data-ttu-id="57544-303">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-303">Requirement</span></span>|<span data-ttu-id="57544-304">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-305">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-306">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-306">1.0</span></span>|
|[<span data-ttu-id="57544-307">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-308">ReadItem</span></span>|
|[<span data-ttu-id="57544-309">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-310">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-310">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="57544-311">(Nullable) Conversation: String</span><span class="sxs-lookup"><span data-stu-id="57544-311">(nullable) conversationId: String</span></span>

<span data-ttu-id="57544-312">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="57544-312">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="57544-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="57544-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="57544-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="57544-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-317">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-317">Type</span></span>

*   <span data-ttu-id="57544-318">String</span><span class="sxs-lookup"><span data-stu-id="57544-318">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-319">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-319">Requirements</span></span>

|<span data-ttu-id="57544-320">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-320">Requirement</span></span>|<span data-ttu-id="57544-321">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-321">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-322">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-322">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-323">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-323">1.0</span></span>|
|[<span data-ttu-id="57544-324">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-324">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-325">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-325">ReadItem</span></span>|
|[<span data-ttu-id="57544-326">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-326">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-327">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-327">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-328">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-328">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="57544-329">dateTimeCreated: data</span><span class="sxs-lookup"><span data-stu-id="57544-329">dateTimeCreated: Date</span></span>

<span data-ttu-id="57544-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="57544-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-332">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-332">Type</span></span>

*   <span data-ttu-id="57544-333">Data</span><span class="sxs-lookup"><span data-stu-id="57544-333">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-334">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-334">Requirements</span></span>

|<span data-ttu-id="57544-335">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-335">Requirement</span></span>|<span data-ttu-id="57544-336">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-337">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-338">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-338">1.0</span></span>|
|[<span data-ttu-id="57544-339">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-340">ReadItem</span></span>|
|[<span data-ttu-id="57544-341">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-342">Read</span><span class="sxs-lookup"><span data-stu-id="57544-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-343">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-343">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="57544-344">dateTimeModified: data</span><span class="sxs-lookup"><span data-stu-id="57544-344">dateTimeModified: Date</span></span>

<span data-ttu-id="57544-345">Obtém a data e a hora em que um item foi alterado pela última vez.</span><span class="sxs-lookup"><span data-stu-id="57544-345">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="57544-346">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="57544-346">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-347">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="57544-347">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-348">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-348">Type</span></span>

*   <span data-ttu-id="57544-349">Data</span><span class="sxs-lookup"><span data-stu-id="57544-349">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-350">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-350">Requirements</span></span>

|<span data-ttu-id="57544-351">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-351">Requirement</span></span>|<span data-ttu-id="57544-352">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-353">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-354">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-354">1.0</span></span>|
|[<span data-ttu-id="57544-355">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-356">ReadItem</span></span>|
|[<span data-ttu-id="57544-357">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-358">Read</span><span class="sxs-lookup"><span data-stu-id="57544-358">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-359">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-359">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="57544-360">fim: data | [Tempo](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="57544-360">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="57544-361">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="57544-361">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="57544-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="57544-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57544-364">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="57544-364">Read mode</span></span>

<span data-ttu-id="57544-365">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="57544-365">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="57544-366">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="57544-366">Compose mode</span></span>

<span data-ttu-id="57544-367">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="57544-367">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="57544-368">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="57544-368">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="57544-369">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="57544-369">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="57544-370">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-370">Type</span></span>

*   <span data-ttu-id="57544-371">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="57544-371">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-372">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-372">Requirements</span></span>

|<span data-ttu-id="57544-373">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-373">Requirement</span></span>|<span data-ttu-id="57544-374">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-375">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-376">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-376">1.0</span></span>|
|[<span data-ttu-id="57544-377">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-378">ReadItem</span></span>|
|[<span data-ttu-id="57544-379">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-380">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-380">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="57544-381">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="57544-381">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="57544-382">Obtém ou define os locais de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-382">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57544-383">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="57544-383">Read mode</span></span>

<span data-ttu-id="57544-384">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que permite que você obtenha o conjunto de locais (cada um representado por um objeto [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associado ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="57544-385">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="57544-385">Compose mode</span></span>

<span data-ttu-id="57544-386">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que fornece métodos para obter, remover ou adicionar locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-386">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-387">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-387">Type</span></span>

*   [<span data-ttu-id="57544-388">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="57544-388">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="57544-389">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-389">Requirements</span></span>

|<span data-ttu-id="57544-390">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-390">Requirement</span></span>|<span data-ttu-id="57544-391">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-391">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-392">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-392">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-393">Visualização</span><span class="sxs-lookup"><span data-stu-id="57544-393">Preview</span></span>|
|[<span data-ttu-id="57544-394">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-395">ReadItem</span></span>|
|[<span data-ttu-id="57544-396">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-396">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-397">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-397">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-398">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-398">Example</span></span>

<span data-ttu-id="57544-399">O exemplo a seguir obtém os locais atuais associados ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-399">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="57544-400">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="57544-400">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="57544-401">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="57544-401">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="57544-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="57544-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-404">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="57544-404">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57544-405">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="57544-405">Read mode</span></span>

<span data-ttu-id="57544-406">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="57544-406">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="57544-407">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="57544-407">Compose mode</span></span>

<span data-ttu-id="57544-408">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="57544-408">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="57544-409">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-409">Type</span></span>

*   <span data-ttu-id="57544-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="57544-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-411">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-411">Requirements</span></span>

|<span data-ttu-id="57544-412">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-412">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="57544-413">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-414">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-414">1.0</span></span>|<span data-ttu-id="57544-415">1.7</span><span class="sxs-lookup"><span data-stu-id="57544-415">1.7</span></span>|
|[<span data-ttu-id="57544-416">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-417">ReadItem</span></span>|<span data-ttu-id="57544-418">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57544-418">ReadWriteItem</span></span>|
|[<span data-ttu-id="57544-419">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-419">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-420">Read</span><span class="sxs-lookup"><span data-stu-id="57544-420">Read</span></span>|<span data-ttu-id="57544-421">Escrever</span><span class="sxs-lookup"><span data-stu-id="57544-421">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="57544-422">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="57544-422">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="57544-423">Obtém ou define cabeçalhos de Internet personalizados em uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="57544-423">Gets or sets custom internet headers on a message.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-424">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-424">Type</span></span>

*   [<span data-ttu-id="57544-425">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="57544-425">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="57544-426">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-426">Requirements</span></span>

|<span data-ttu-id="57544-427">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-427">Requirement</span></span>|<span data-ttu-id="57544-428">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-429">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-430">Visualização</span><span class="sxs-lookup"><span data-stu-id="57544-430">Preview</span></span>|
|[<span data-ttu-id="57544-431">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-432">ReadItem</span></span>|
|[<span data-ttu-id="57544-433">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-434">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-435">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-435">Example</span></span>

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

#### <a name="internetmessageid-string"></a><span data-ttu-id="57544-436">internetMessageId: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="57544-436">internetMessageId: String</span></span>

<span data-ttu-id="57544-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="57544-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-439">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-439">Type</span></span>

*   <span data-ttu-id="57544-440">String</span><span class="sxs-lookup"><span data-stu-id="57544-440">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-441">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-441">Requirements</span></span>

|<span data-ttu-id="57544-442">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-442">Requirement</span></span>|<span data-ttu-id="57544-443">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-444">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-445">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-445">1.0</span></span>|
|[<span data-ttu-id="57544-446">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-447">ReadItem</span></span>|
|[<span data-ttu-id="57544-448">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-449">Read</span><span class="sxs-lookup"><span data-stu-id="57544-449">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-450">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-450">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="57544-451">doclass: String</span><span class="sxs-lookup"><span data-stu-id="57544-451">itemClass: String</span></span>

<span data-ttu-id="57544-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="57544-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="57544-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="57544-456">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-456">Type</span></span>|<span data-ttu-id="57544-457">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-457">Description</span></span>|<span data-ttu-id="57544-458">classe de item</span><span class="sxs-lookup"><span data-stu-id="57544-458">item class</span></span>|
|---|---|---|
|<span data-ttu-id="57544-459">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="57544-459">Appointment items</span></span>|<span data-ttu-id="57544-460">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="57544-460">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="57544-461">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="57544-461">Message items</span></span>|<span data-ttu-id="57544-462">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="57544-462">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="57544-463">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="57544-463">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-464">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-464">Type</span></span>

*   <span data-ttu-id="57544-465">String</span><span class="sxs-lookup"><span data-stu-id="57544-465">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-466">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-466">Requirements</span></span>

|<span data-ttu-id="57544-467">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-467">Requirement</span></span>|<span data-ttu-id="57544-468">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-469">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-470">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-470">1.0</span></span>|
|[<span data-ttu-id="57544-471">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-472">ReadItem</span></span>|
|[<span data-ttu-id="57544-473">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-474">Read</span><span class="sxs-lookup"><span data-stu-id="57544-474">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-475">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-475">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="57544-476">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="57544-476">(nullable) itemId: String</span></span>

<span data-ttu-id="57544-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="57544-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-479">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="57544-479">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="57544-480">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="57544-480">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="57544-481">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="57544-481">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="57544-482">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="57544-482">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="57544-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-485">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-485">Type</span></span>

*   <span data-ttu-id="57544-486">String</span><span class="sxs-lookup"><span data-stu-id="57544-486">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-487">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-487">Requirements</span></span>

|<span data-ttu-id="57544-488">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-488">Requirement</span></span>|<span data-ttu-id="57544-489">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-490">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-490">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-491">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-491">1.0</span></span>|
|[<span data-ttu-id="57544-492">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-492">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-493">ReadItem</span></span>|
|[<span data-ttu-id="57544-494">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-494">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-495">Read</span><span class="sxs-lookup"><span data-stu-id="57544-495">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-496">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-496">Example</span></span>

<span data-ttu-id="57544-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="57544-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="57544-499">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="57544-499">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="57544-500">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="57544-500">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="57544-501">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-501">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-502">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-502">Type</span></span>

*   [<span data-ttu-id="57544-503">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="57544-503">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="57544-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-504">Requirements</span></span>

|<span data-ttu-id="57544-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-505">Requirement</span></span>|<span data-ttu-id="57544-506">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-507">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-508">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-508">1.0</span></span>|
|[<span data-ttu-id="57544-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-510">ReadItem</span></span>|
|[<span data-ttu-id="57544-511">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-512">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-512">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-513">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-513">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="57544-514">local: cadeia de caracteres | [Local](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="57544-514">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="57544-515">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-515">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57544-516">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="57544-516">Read mode</span></span>

<span data-ttu-id="57544-517">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-517">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="57544-518">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="57544-518">Compose mode</span></span>

<span data-ttu-id="57544-519">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-519">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="57544-520">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-520">Type</span></span>

*   <span data-ttu-id="57544-521">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="57544-521">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-522">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-522">Requirements</span></span>

|<span data-ttu-id="57544-523">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-523">Requirement</span></span>|<span data-ttu-id="57544-524">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-525">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-526">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-526">1.0</span></span>|
|[<span data-ttu-id="57544-527">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-527">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-528">ReadItem</span></span>|
|[<span data-ttu-id="57544-529">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-530">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-530">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="57544-531">normalizedSubject: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="57544-531">normalizedSubject: String</span></span>

<span data-ttu-id="57544-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="57544-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="57544-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="57544-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-536">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-536">Type</span></span>

*   <span data-ttu-id="57544-537">String</span><span class="sxs-lookup"><span data-stu-id="57544-537">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-538">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-538">Requirements</span></span>

|<span data-ttu-id="57544-539">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-539">Requirement</span></span>|<span data-ttu-id="57544-540">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-541">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-542">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-542">1.0</span></span>|
|[<span data-ttu-id="57544-543">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-544">ReadItem</span></span>|
|[<span data-ttu-id="57544-545">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-546">Read</span><span class="sxs-lookup"><span data-stu-id="57544-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-547">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-547">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="57544-548">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="57544-548">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="57544-549">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="57544-549">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-550">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-550">Type</span></span>

*   [<span data-ttu-id="57544-551">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="57544-551">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="57544-552">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-552">Requirements</span></span>

|<span data-ttu-id="57544-553">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-553">Requirement</span></span>|<span data-ttu-id="57544-554">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-554">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-555">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-555">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-556">1.3</span><span class="sxs-lookup"><span data-stu-id="57544-556">1.3</span></span>|
|[<span data-ttu-id="57544-557">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-557">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-558">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-558">ReadItem</span></span>|
|[<span data-ttu-id="57544-559">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-559">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-560">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-560">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-561">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-561">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="57544-562">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[](/javascript/api/outlook/office.recipients) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="57544-562">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="57544-563">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="57544-563">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="57544-564">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="57544-564">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57544-565">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="57544-565">Read mode</span></span>

<span data-ttu-id="57544-566">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="57544-566">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="57544-567">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="57544-567">Compose mode</span></span>

<span data-ttu-id="57544-568">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="57544-568">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="57544-569">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-569">Type</span></span>

*   <span data-ttu-id="57544-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="57544-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-571">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-571">Requirements</span></span>

|<span data-ttu-id="57544-572">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-572">Requirement</span></span>|<span data-ttu-id="57544-573">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-573">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-574">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-575">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-575">1.0</span></span>|
|[<span data-ttu-id="57544-576">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-577">ReadItem</span></span>|
|[<span data-ttu-id="57544-578">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-579">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-579">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="57544-580">organizador: [](/javascript/api/outlook/office.emailaddressdetails)|[organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="57544-580">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="57544-581">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="57544-581">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57544-582">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="57544-582">Read mode</span></span>

<span data-ttu-id="57544-583">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="57544-583">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="57544-584">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="57544-584">Compose mode</span></span>

<span data-ttu-id="57544-585">A `organizer` propriedade retorna um objeto [organizador](/javascript/api/outlook/office.organizer) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="57544-585">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="57544-586">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-586">Type</span></span>

*   <span data-ttu-id="57544-587">[](/javascript/api/outlook/office.emailaddressdetails) | [Organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="57544-587">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-588">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-588">Requirements</span></span>

|<span data-ttu-id="57544-589">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-589">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="57544-590">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-591">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-591">1.0</span></span>|<span data-ttu-id="57544-592">1.7</span><span class="sxs-lookup"><span data-stu-id="57544-592">1.7</span></span>|
|[<span data-ttu-id="57544-593">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-594">ReadItem</span></span>|<span data-ttu-id="57544-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57544-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="57544-596">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-597">Read</span><span class="sxs-lookup"><span data-stu-id="57544-597">Read</span></span>|<span data-ttu-id="57544-598">Escrever</span><span class="sxs-lookup"><span data-stu-id="57544-598">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="57544-599">(anulável) recorrência: [recorrência](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="57544-599">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="57544-600">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-600">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="57544-601">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="57544-601">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="57544-602">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-602">Read and compose modes for appointment items.</span></span> <span data-ttu-id="57544-603">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="57544-603">Read mode for meeting request items.</span></span>

<span data-ttu-id="57544-604">A `recurrence` propriedade retorna um objeto de [recorrência](/javascript/api/outlook/office.recurrence) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="57544-604">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="57544-605">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="57544-605">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="57544-606">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="57544-606">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="57544-607">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="57544-607">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="57544-608">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="57544-608">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57544-609">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="57544-609">Read mode</span></span>

<span data-ttu-id="57544-610">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-610">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="57544-611">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="57544-611">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="57544-612">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="57544-612">Compose mode</span></span>

<span data-ttu-id="57544-613">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-613">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="57544-614">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="57544-614">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="57544-615">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-615">Type</span></span>

* [<span data-ttu-id="57544-616">Recorrência</span><span class="sxs-lookup"><span data-stu-id="57544-616">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="57544-617">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-617">Requirement</span></span>|<span data-ttu-id="57544-618">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-619">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-620">1.7</span><span class="sxs-lookup"><span data-stu-id="57544-620">1.7</span></span>|
|[<span data-ttu-id="57544-621">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-622">ReadItem</span></span>|
|[<span data-ttu-id="57544-623">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-624">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-624">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="57544-625">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[](/javascript/api/outlook/office.recipients) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="57544-625">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="57544-626">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="57544-626">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="57544-627">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="57544-627">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57544-628">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="57544-628">Read mode</span></span>

<span data-ttu-id="57544-629">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="57544-629">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="57544-630">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="57544-630">Compose mode</span></span>

<span data-ttu-id="57544-631">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="57544-631">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="57544-632">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-632">Type</span></span>

*   <span data-ttu-id="57544-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="57544-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-634">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-634">Requirements</span></span>

|<span data-ttu-id="57544-635">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-635">Requirement</span></span>|<span data-ttu-id="57544-636">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-637">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-638">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-638">1.0</span></span>|
|[<span data-ttu-id="57544-639">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-640">ReadItem</span></span>|
|[<span data-ttu-id="57544-641">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-642">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-642">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="57544-643">remetente: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="57544-643">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="57544-p128">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="57544-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="57544-p129">As propriedades [`from`](#from-emailaddressdetailsfrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="57544-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-648">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="57544-648">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-649">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-649">Type</span></span>

*   [<span data-ttu-id="57544-650">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="57544-650">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="57544-651">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-651">Requirements</span></span>

|<span data-ttu-id="57544-652">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-652">Requirement</span></span>|<span data-ttu-id="57544-653">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-654">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-655">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-655">1.0</span></span>|
|[<span data-ttu-id="57544-656">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-657">ReadItem</span></span>|
|[<span data-ttu-id="57544-658">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-659">Read</span><span class="sxs-lookup"><span data-stu-id="57544-659">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-660">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-660">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="57544-661">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="57544-661">(nullable) seriesId: String</span></span>

<span data-ttu-id="57544-662">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="57544-662">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="57544-663">No Outlook na Web e clientes de desktop, o `seriesId` retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="57544-663">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="57544-664">No entanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="57544-664">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-665">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="57544-665">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="57544-666">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="57544-666">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="57544-667">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="57544-667">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="57544-668">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="57544-668">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="57544-669">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="57544-669">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="57544-670">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-670">Type</span></span>

* <span data-ttu-id="57544-671">String</span><span class="sxs-lookup"><span data-stu-id="57544-671">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-672">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-672">Requirements</span></span>

|<span data-ttu-id="57544-673">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-673">Requirement</span></span>|<span data-ttu-id="57544-674">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-674">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-675">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-675">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-676">1.7</span><span class="sxs-lookup"><span data-stu-id="57544-676">1.7</span></span>|
|[<span data-ttu-id="57544-677">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-677">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-678">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-678">ReadItem</span></span>|
|[<span data-ttu-id="57544-679">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-679">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-680">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-680">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-681">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-681">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="57544-682">Início: data | [Tempo](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="57544-682">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="57544-683">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="57544-683">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="57544-p132">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="57544-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57544-686">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="57544-686">Read mode</span></span>

<span data-ttu-id="57544-687">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="57544-687">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="57544-688">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="57544-688">Compose mode</span></span>

<span data-ttu-id="57544-689">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="57544-689">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="57544-690">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="57544-690">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="57544-691">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="57544-691">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="57544-692">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-692">Type</span></span>

*   <span data-ttu-id="57544-693">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="57544-693">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-694">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-694">Requirements</span></span>

|<span data-ttu-id="57544-695">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-695">Requirement</span></span>|<span data-ttu-id="57544-696">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-697">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-698">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-698">1.0</span></span>|
|[<span data-ttu-id="57544-699">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-700">ReadItem</span></span>|
|[<span data-ttu-id="57544-701">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-702">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-702">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="57544-703">subject: cadeia de caracteres | [Assunto](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="57544-703">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="57544-704">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="57544-704">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="57544-705">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="57544-705">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57544-706">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="57544-706">Read mode</span></span>

<span data-ttu-id="57544-p133">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="57544-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="57544-709">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="57544-709">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="57544-710">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="57544-710">Compose mode</span></span>
<span data-ttu-id="57544-711">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="57544-711">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="57544-712">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-712">Type</span></span>

*   <span data-ttu-id="57544-713">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="57544-713">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-714">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-714">Requirements</span></span>

|<span data-ttu-id="57544-715">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-715">Requirement</span></span>|<span data-ttu-id="57544-716">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-716">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-717">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-717">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-718">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-718">1.0</span></span>|
|[<span data-ttu-id="57544-719">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-719">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-720">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-720">ReadItem</span></span>|
|[<span data-ttu-id="57544-721">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-721">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-722">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-722">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="57544-723">para: Array. <[](/javascript/api/outlook/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook/office.recipients) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="57544-723">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="57544-724">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="57544-724">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="57544-725">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="57544-725">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="57544-726">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="57544-726">Read mode</span></span>

<span data-ttu-id="57544-p135">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="57544-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="57544-729">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="57544-729">Compose mode</span></span>

<span data-ttu-id="57544-730">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="57544-730">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="57544-731">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-731">Type</span></span>

*   <span data-ttu-id="57544-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="57544-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-733">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-733">Requirements</span></span>

|<span data-ttu-id="57544-734">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-734">Requirement</span></span>|<span data-ttu-id="57544-735">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-735">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-736">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-736">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-737">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-737">1.0</span></span>|
|[<span data-ttu-id="57544-738">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-738">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-739">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-739">ReadItem</span></span>|
|[<span data-ttu-id="57544-740">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-740">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-741">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-741">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="57544-742">Métodos</span><span class="sxs-lookup"><span data-stu-id="57544-742">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="57544-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="57544-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="57544-744">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="57544-744">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="57544-745">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="57544-745">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="57544-746">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="57544-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-747">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-747">Parameters</span></span>
|<span data-ttu-id="57544-748">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-748">Name</span></span>|<span data-ttu-id="57544-749">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-749">Type</span></span>|<span data-ttu-id="57544-750">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-750">Attributes</span></span>|<span data-ttu-id="57544-751">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-751">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="57544-752">String</span><span class="sxs-lookup"><span data-stu-id="57544-752">String</span></span>||<span data-ttu-id="57544-p136">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="57544-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="57544-755">String</span><span class="sxs-lookup"><span data-stu-id="57544-755">String</span></span>||<span data-ttu-id="57544-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="57544-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="57544-758">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-758">Object</span></span>|<span data-ttu-id="57544-759">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-759">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-760">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57544-761">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-761">Object</span></span>|<span data-ttu-id="57544-762">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-762">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-763">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="57544-764">Booliano</span><span class="sxs-lookup"><span data-stu-id="57544-764">Boolean</span></span>|<span data-ttu-id="57544-765">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-765">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-766">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="57544-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="57544-767">function</span><span class="sxs-lookup"><span data-stu-id="57544-767">function</span></span>|<span data-ttu-id="57544-768">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-768">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-769">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="57544-770">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="57544-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="57544-771">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="57544-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="57544-772">Erros</span><span class="sxs-lookup"><span data-stu-id="57544-772">Errors</span></span>

|<span data-ttu-id="57544-773">Código de erro</span><span class="sxs-lookup"><span data-stu-id="57544-773">Error code</span></span>|<span data-ttu-id="57544-774">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="57544-775">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="57544-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="57544-776">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="57544-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="57544-777">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="57544-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-778">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-778">Requirements</span></span>

|<span data-ttu-id="57544-779">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-779">Requirement</span></span>|<span data-ttu-id="57544-780">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-781">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-782">1.1</span><span class="sxs-lookup"><span data-stu-id="57544-782">1.1</span></span>|
|[<span data-ttu-id="57544-783">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57544-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="57544-785">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-786">Escrever</span><span class="sxs-lookup"><span data-stu-id="57544-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="57544-787">Exemplos</span><span class="sxs-lookup"><span data-stu-id="57544-787">Examples</span></span>

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

<span data-ttu-id="57544-788">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="57544-788">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="57544-789">addFileAttachmentFromBase64Async (base64file, AttachmentName, [Options], [callback])</span><span class="sxs-lookup"><span data-stu-id="57544-789">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="57544-790">Adiciona um arquivo da codificação Base64 a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="57544-790">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="57544-791">O `addFileAttachmentFromBase64Async` método carrega o arquivo da codificação Base64 e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="57544-791">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="57544-792">Esse método retorna o identificador de anexo no objeto AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="57544-792">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="57544-793">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="57544-793">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-794">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-794">Parameters</span></span>

|<span data-ttu-id="57544-795">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-795">Name</span></span>|<span data-ttu-id="57544-796">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-796">Type</span></span>|<span data-ttu-id="57544-797">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-797">Attributes</span></span>|<span data-ttu-id="57544-798">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-798">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="57544-799">String</span><span class="sxs-lookup"><span data-stu-id="57544-799">String</span></span>||<span data-ttu-id="57544-800">O conteúdo codificado em Base64 de uma imagem ou arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="57544-800">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="57544-801">String</span><span class="sxs-lookup"><span data-stu-id="57544-801">String</span></span>||<span data-ttu-id="57544-p139">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="57544-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="57544-804">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-804">Object</span></span>|<span data-ttu-id="57544-805">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-805">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-806">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-806">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57544-807">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-807">Object</span></span>|<span data-ttu-id="57544-808">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-808">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-809">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-809">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="57544-810">Booliano</span><span class="sxs-lookup"><span data-stu-id="57544-810">Boolean</span></span>|<span data-ttu-id="57544-811">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-811">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-812">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="57544-812">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="57544-813">function</span><span class="sxs-lookup"><span data-stu-id="57544-813">function</span></span>|<span data-ttu-id="57544-814">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-814">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-815">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-815">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="57544-816">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="57544-816">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="57544-817">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="57544-817">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="57544-818">Erros</span><span class="sxs-lookup"><span data-stu-id="57544-818">Errors</span></span>

|<span data-ttu-id="57544-819">Código de erro</span><span class="sxs-lookup"><span data-stu-id="57544-819">Error code</span></span>|<span data-ttu-id="57544-820">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-820">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="57544-821">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="57544-821">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="57544-822">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="57544-822">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="57544-823">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="57544-823">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-824">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-824">Requirements</span></span>

|<span data-ttu-id="57544-825">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-825">Requirement</span></span>|<span data-ttu-id="57544-826">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-827">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-828">Visualização</span><span class="sxs-lookup"><span data-stu-id="57544-828">Preview</span></span>|
|[<span data-ttu-id="57544-829">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-830">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57544-830">ReadWriteItem</span></span>|
|[<span data-ttu-id="57544-831">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-832">Escrever</span><span class="sxs-lookup"><span data-stu-id="57544-832">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="57544-833">Exemplos</span><span class="sxs-lookup"><span data-stu-id="57544-833">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="57544-834">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="57544-834">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="57544-835">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="57544-835">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="57544-836">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="57544-836">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-837">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-837">Parameters</span></span>

| <span data-ttu-id="57544-838">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-838">Name</span></span> | <span data-ttu-id="57544-839">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-839">Type</span></span> | <span data-ttu-id="57544-840">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-840">Attributes</span></span> | <span data-ttu-id="57544-841">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-841">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="57544-842">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="57544-842">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="57544-843">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="57544-843">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="57544-844">Função</span><span class="sxs-lookup"><span data-stu-id="57544-844">Function</span></span> || <span data-ttu-id="57544-p140">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="57544-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="57544-848">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-848">Object</span></span> | <span data-ttu-id="57544-849">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-849">&lt;optional&gt;</span></span> | <span data-ttu-id="57544-850">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-850">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="57544-851">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-851">Object</span></span> | <span data-ttu-id="57544-852">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-852">&lt;optional&gt;</span></span> | <span data-ttu-id="57544-853">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-853">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="57544-854">function</span><span class="sxs-lookup"><span data-stu-id="57544-854">function</span></span>| <span data-ttu-id="57544-855">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-855">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-856">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-857">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-857">Requirements</span></span>

|<span data-ttu-id="57544-858">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-858">Requirement</span></span>| <span data-ttu-id="57544-859">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-859">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-860">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57544-861">1.7</span><span class="sxs-lookup"><span data-stu-id="57544-861">1.7</span></span> |
|[<span data-ttu-id="57544-862">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="57544-863">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-863">ReadItem</span></span> |
|[<span data-ttu-id="57544-864">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57544-865">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-865">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="57544-866">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-866">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="57544-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="57544-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="57544-868">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-868">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="57544-p141">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="57544-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="57544-872">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="57544-872">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="57544-873">Se o suplemento do Office estiver em execução no Outlook na Web, o `addItemAttachmentAsync` método poderá anexar itens a itens diferentes do item que você está editando; no entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="57544-873">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-874">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-874">Parameters</span></span>

|<span data-ttu-id="57544-875">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-875">Name</span></span>|<span data-ttu-id="57544-876">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-876">Type</span></span>|<span data-ttu-id="57544-877">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-877">Attributes</span></span>|<span data-ttu-id="57544-878">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-878">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="57544-879">String</span><span class="sxs-lookup"><span data-stu-id="57544-879">String</span></span>||<span data-ttu-id="57544-p142">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="57544-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="57544-882">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="57544-882">String</span></span>||<span data-ttu-id="57544-883">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="57544-883">The subject of the item to be attached.</span></span> <span data-ttu-id="57544-884">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="57544-884">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="57544-885">Object</span><span class="sxs-lookup"><span data-stu-id="57544-885">Object</span></span>|<span data-ttu-id="57544-886">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-886">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-887">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-887">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57544-888">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-888">Object</span></span>|<span data-ttu-id="57544-889">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-889">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-890">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-890">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57544-891">function</span><span class="sxs-lookup"><span data-stu-id="57544-891">function</span></span>|<span data-ttu-id="57544-892">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-892">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-893">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-893">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="57544-894">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="57544-894">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="57544-895">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="57544-895">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="57544-896">Erros</span><span class="sxs-lookup"><span data-stu-id="57544-896">Errors</span></span>

|<span data-ttu-id="57544-897">Código de erro</span><span class="sxs-lookup"><span data-stu-id="57544-897">Error code</span></span>|<span data-ttu-id="57544-898">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-898">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="57544-899">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="57544-899">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-900">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-900">Requirements</span></span>

|<span data-ttu-id="57544-901">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-901">Requirement</span></span>|<span data-ttu-id="57544-902">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-903">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-904">1.1</span><span class="sxs-lookup"><span data-stu-id="57544-904">1.1</span></span>|
|[<span data-ttu-id="57544-905">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57544-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="57544-907">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-908">Escrever</span><span class="sxs-lookup"><span data-stu-id="57544-908">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-909">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-909">Example</span></span>

<span data-ttu-id="57544-910">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="57544-910">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="57544-911">close()</span><span class="sxs-lookup"><span data-stu-id="57544-911">close()</span></span>

<span data-ttu-id="57544-912">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="57544-912">Closes the current item that is being composed.</span></span>

<span data-ttu-id="57544-p144">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="57544-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-915">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="57544-915">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="57544-916">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="57544-916">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-917">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-917">Requirements</span></span>

|<span data-ttu-id="57544-918">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-918">Requirement</span></span>|<span data-ttu-id="57544-919">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-919">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-920">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-920">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-921">1.3</span><span class="sxs-lookup"><span data-stu-id="57544-921">1.3</span></span>|
|[<span data-ttu-id="57544-922">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-922">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-923">Restrito</span><span class="sxs-lookup"><span data-stu-id="57544-923">Restricted</span></span>|
|[<span data-ttu-id="57544-924">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-924">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-925">Escrever</span><span class="sxs-lookup"><span data-stu-id="57544-925">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="57544-926">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="57544-926">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="57544-927">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="57544-927">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-928">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="57544-928">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="57544-929">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="57544-929">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="57544-930">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="57544-930">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="57544-931">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="57544-931">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="57544-932">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="57544-932">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="57544-933">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="57544-933">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-934">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-934">Parameters</span></span>

|<span data-ttu-id="57544-935">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-935">Name</span></span>|<span data-ttu-id="57544-936">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-936">Type</span></span>|<span data-ttu-id="57544-937">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-937">Attributes</span></span>|<span data-ttu-id="57544-938">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-938">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="57544-939">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="57544-939">String &#124; Object</span></span>||<span data-ttu-id="57544-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="57544-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="57544-942">**OU**</span><span class="sxs-lookup"><span data-stu-id="57544-942">**OR**</span></span><br/><span data-ttu-id="57544-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="57544-945">String</span><span class="sxs-lookup"><span data-stu-id="57544-945">String</span></span>|<span data-ttu-id="57544-946">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-946">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="57544-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="57544-949">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-949">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="57544-950">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-950">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-951">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="57544-951">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="57544-952">String</span><span class="sxs-lookup"><span data-stu-id="57544-952">String</span></span>||<span data-ttu-id="57544-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="57544-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="57544-955">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="57544-955">String</span></span>||<span data-ttu-id="57544-956">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="57544-956">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="57544-957">String</span><span class="sxs-lookup"><span data-stu-id="57544-957">String</span></span>||<span data-ttu-id="57544-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="57544-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="57544-960">Booliano</span><span class="sxs-lookup"><span data-stu-id="57544-960">Boolean</span></span>||<span data-ttu-id="57544-p151">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="57544-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="57544-963">String</span><span class="sxs-lookup"><span data-stu-id="57544-963">String</span></span>||<span data-ttu-id="57544-p152">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="57544-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="57544-967">function</span><span class="sxs-lookup"><span data-stu-id="57544-967">function</span></span>|<span data-ttu-id="57544-968">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-968">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-969">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-969">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-970">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-970">Requirements</span></span>

|<span data-ttu-id="57544-971">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-971">Requirement</span></span>|<span data-ttu-id="57544-972">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-972">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-973">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-973">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-974">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-974">1.0</span></span>|
|[<span data-ttu-id="57544-975">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-975">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-976">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-976">ReadItem</span></span>|
|[<span data-ttu-id="57544-977">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-977">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-978">Read</span><span class="sxs-lookup"><span data-stu-id="57544-978">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="57544-979">Exemplos</span><span class="sxs-lookup"><span data-stu-id="57544-979">Examples</span></span>

<span data-ttu-id="57544-980">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="57544-980">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="57544-981">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="57544-981">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="57544-982">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="57544-982">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="57544-983">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="57544-983">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="57544-984">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="57544-984">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="57544-985">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-985">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="57544-986">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="57544-986">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="57544-987">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="57544-987">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-988">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="57544-988">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="57544-989">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="57544-989">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="57544-990">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="57544-990">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="57544-991">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="57544-991">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="57544-992">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="57544-992">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="57544-993">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="57544-993">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-994">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-994">Parameters</span></span>

|<span data-ttu-id="57544-995">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-995">Name</span></span>|<span data-ttu-id="57544-996">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-996">Type</span></span>|<span data-ttu-id="57544-997">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-997">Attributes</span></span>|<span data-ttu-id="57544-998">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-998">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="57544-999">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="57544-999">String &#124; Object</span></span>||<span data-ttu-id="57544-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="57544-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="57544-1002">**OU**</span><span class="sxs-lookup"><span data-stu-id="57544-1002">**OR**</span></span><br/><span data-ttu-id="57544-p155">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="57544-1005">String</span><span class="sxs-lookup"><span data-stu-id="57544-1005">String</span></span>|<span data-ttu-id="57544-1006">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-p156">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="57544-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="57544-1009">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1009">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="57544-1010">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1011">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="57544-1011">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="57544-1012">String</span><span class="sxs-lookup"><span data-stu-id="57544-1012">String</span></span>||<span data-ttu-id="57544-p157">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="57544-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="57544-1015">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="57544-1015">String</span></span>||<span data-ttu-id="57544-1016">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="57544-1016">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="57544-1017">String</span><span class="sxs-lookup"><span data-stu-id="57544-1017">String</span></span>||<span data-ttu-id="57544-p158">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="57544-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="57544-1020">Booliano</span><span class="sxs-lookup"><span data-stu-id="57544-1020">Boolean</span></span>||<span data-ttu-id="57544-p159">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="57544-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="57544-1023">String</span><span class="sxs-lookup"><span data-stu-id="57544-1023">String</span></span>||<span data-ttu-id="57544-p160">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="57544-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="57544-1027">function</span><span class="sxs-lookup"><span data-stu-id="57544-1027">function</span></span>|<span data-ttu-id="57544-1028">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1028">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1029">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-1029">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1030">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1030">Requirements</span></span>

|<span data-ttu-id="57544-1031">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1031">Requirement</span></span>|<span data-ttu-id="57544-1032">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1032">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1033">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1033">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1034">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-1034">1.0</span></span>|
|[<span data-ttu-id="57544-1035">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1035">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1036">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1036">ReadItem</span></span>|
|[<span data-ttu-id="57544-1037">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1037">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1038">Read</span><span class="sxs-lookup"><span data-stu-id="57544-1038">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="57544-1039">Exemplos</span><span class="sxs-lookup"><span data-stu-id="57544-1039">Examples</span></span>

<span data-ttu-id="57544-1040">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="57544-1040">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="57544-1041">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="57544-1041">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="57544-1042">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="57544-1042">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="57544-1043">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="57544-1043">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="57544-1044">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="57544-1044">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="57544-1045">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1045">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="57544-1046">getAttachmentContentAsync (attachmentid, [opções], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="57544-1046">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="57544-1047">Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um `AttachmentContent` objeto.</span><span class="sxs-lookup"><span data-stu-id="57544-1047">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="57544-1048">O `getAttachmentContentAsync` método obtém o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="57544-1048">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="57544-1049">Como prática recomendada, você deve usar o identificador para recuperar um anexo na mesma sessão em que o attachmentIds foi recuperado com a `getAttachmentsAsync` chamada ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="57544-1049">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="57544-1050">No Outlook na Web e dispositivos móveis, o identificador de anexo é válido somente dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="57544-1050">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="57544-1051">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="57544-1051">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1052">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1052">Parameters</span></span>

|<span data-ttu-id="57544-1053">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1053">Name</span></span>|<span data-ttu-id="57544-1054">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1054">Type</span></span>|<span data-ttu-id="57544-1055">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-1055">Attributes</span></span>|<span data-ttu-id="57544-1056">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1056">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="57544-1057">String</span><span class="sxs-lookup"><span data-stu-id="57544-1057">String</span></span>||<span data-ttu-id="57544-1058">O identificador do anexo que você deseja obter.</span><span class="sxs-lookup"><span data-stu-id="57544-1058">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="57544-1059">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1059">Object</span></span>|<span data-ttu-id="57544-1060">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1060">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1061">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-1061">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57544-1062">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1062">Object</span></span>|<span data-ttu-id="57544-1063">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1064">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1064">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57544-1065">function</span><span class="sxs-lookup"><span data-stu-id="57544-1065">function</span></span>|<span data-ttu-id="57544-1066">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1067">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-1067">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1068">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1068">Requirements</span></span>

|<span data-ttu-id="57544-1069">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1069">Requirement</span></span>|<span data-ttu-id="57544-1070">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1070">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1071">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1071">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1072">Visualização</span><span class="sxs-lookup"><span data-stu-id="57544-1072">Preview</span></span>|
|[<span data-ttu-id="57544-1073">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1073">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1074">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1074">ReadItem</span></span>|
|[<span data-ttu-id="57544-1075">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-1075">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1076">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-1076">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57544-1077">Retorna:</span><span class="sxs-lookup"><span data-stu-id="57544-1077">Returns:</span></span>

<span data-ttu-id="57544-1078">Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="57544-1078">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="57544-1079">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1079">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="57544-1080">getAttachmentsAsync ([Options], [callback]) → array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="57544-1080">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="57544-1081">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="57544-1081">Gets the item's attachments as an array.</span></span> <span data-ttu-id="57544-1082">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="57544-1082">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1083">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1083">Parameters</span></span>

|<span data-ttu-id="57544-1084">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1084">Name</span></span>|<span data-ttu-id="57544-1085">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1085">Type</span></span>|<span data-ttu-id="57544-1086">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-1086">Attributes</span></span>|<span data-ttu-id="57544-1087">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1087">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="57544-1088">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1088">Object</span></span>|<span data-ttu-id="57544-1089">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1089">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1090">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-1090">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57544-1091">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1091">Object</span></span>|<span data-ttu-id="57544-1092">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1093">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1093">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57544-1094">function</span><span class="sxs-lookup"><span data-stu-id="57544-1094">function</span></span>|<span data-ttu-id="57544-1095">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1096">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-1096">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1097">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1097">Requirements</span></span>

|<span data-ttu-id="57544-1098">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1098">Requirement</span></span>|<span data-ttu-id="57544-1099">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1099">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1100">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1100">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1101">Visualização</span><span class="sxs-lookup"><span data-stu-id="57544-1101">Preview</span></span>|
|[<span data-ttu-id="57544-1102">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1102">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1103">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1103">ReadItem</span></span>|
|[<span data-ttu-id="57544-1104">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1104">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1105">Escrever</span><span class="sxs-lookup"><span data-stu-id="57544-1105">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="57544-1106">Retorna:</span><span class="sxs-lookup"><span data-stu-id="57544-1106">Returns:</span></span>

<span data-ttu-id="57544-1107">Tipo: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="57544-1107">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="57544-1108">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1108">Example</span></span>

<span data-ttu-id="57544-1109">O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="57544-1109">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="57544-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="57544-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="57544-1111">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="57544-1111">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-1112">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="57544-1112">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-1113">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1113">Requirements</span></span>

|<span data-ttu-id="57544-1114">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1114">Requirement</span></span>|<span data-ttu-id="57544-1115">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1115">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1116">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1116">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1117">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-1117">1.0</span></span>|
|[<span data-ttu-id="57544-1118">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1118">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1119">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1119">ReadItem</span></span>|
|[<span data-ttu-id="57544-1120">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1120">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1121">Read</span><span class="sxs-lookup"><span data-stu-id="57544-1121">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57544-1122">Retorna:</span><span class="sxs-lookup"><span data-stu-id="57544-1122">Returns:</span></span>

<span data-ttu-id="57544-1123">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="57544-1123">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="57544-1124">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1124">Example</span></span>

<span data-ttu-id="57544-1125">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="57544-1125">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="57544-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="57544-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="57544-1127">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="57544-1127">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-1128">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="57544-1128">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1129">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1129">Parameters</span></span>

|<span data-ttu-id="57544-1130">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1130">Name</span></span>|<span data-ttu-id="57544-1131">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1131">Type</span></span>|<span data-ttu-id="57544-1132">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1132">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="57544-1133">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="57544-1133">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="57544-1134">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="57544-1134">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1135">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1135">Requirements</span></span>

|<span data-ttu-id="57544-1136">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1136">Requirement</span></span>|<span data-ttu-id="57544-1137">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1138">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1139">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-1139">1.0</span></span>|
|[<span data-ttu-id="57544-1140">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1141">Restrito</span><span class="sxs-lookup"><span data-stu-id="57544-1141">Restricted</span></span>|
|[<span data-ttu-id="57544-1142">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1143">Read</span><span class="sxs-lookup"><span data-stu-id="57544-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57544-1144">Retorna:</span><span class="sxs-lookup"><span data-stu-id="57544-1144">Returns:</span></span>

<span data-ttu-id="57544-1145">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="57544-1145">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="57544-1146">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="57544-1146">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="57544-1147">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="57544-1147">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="57544-1148">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-1148">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="57544-1149">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="57544-1149">Value of `entityType`</span></span>|<span data-ttu-id="57544-1150">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="57544-1150">Type of objects in returned array</span></span>|<span data-ttu-id="57544-1151">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="57544-1151">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="57544-1152">String</span><span class="sxs-lookup"><span data-stu-id="57544-1152">String</span></span>|<span data-ttu-id="57544-1153">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="57544-1153">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="57544-1154">Contato</span><span class="sxs-lookup"><span data-stu-id="57544-1154">Contact</span></span>|<span data-ttu-id="57544-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="57544-1155">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="57544-1156">String</span><span class="sxs-lookup"><span data-stu-id="57544-1156">String</span></span>|<span data-ttu-id="57544-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="57544-1157">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="57544-1158">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="57544-1158">MeetingSuggestion</span></span>|<span data-ttu-id="57544-1159">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="57544-1159">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="57544-1160">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="57544-1160">PhoneNumber</span></span>|<span data-ttu-id="57544-1161">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="57544-1161">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="57544-1162">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="57544-1162">TaskSuggestion</span></span>|<span data-ttu-id="57544-1163">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="57544-1163">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="57544-1164">String</span><span class="sxs-lookup"><span data-stu-id="57544-1164">String</span></span>|<span data-ttu-id="57544-1165">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="57544-1165">**Restricted**</span></span>|

<span data-ttu-id="57544-1166">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="57544-1166">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="57544-1167">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1167">Example</span></span>

<span data-ttu-id="57544-1168">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="57544-1168">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="57544-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="57544-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="57544-1170">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="57544-1170">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-1171">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="57544-1171">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="57544-1172">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="57544-1172">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1173">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1173">Parameters</span></span>

|<span data-ttu-id="57544-1174">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1174">Name</span></span>|<span data-ttu-id="57544-1175">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1175">Type</span></span>|<span data-ttu-id="57544-1176">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1176">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="57544-1177">String</span><span class="sxs-lookup"><span data-stu-id="57544-1177">String</span></span>|<span data-ttu-id="57544-1178">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="57544-1178">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1179">Requirements</span></span>

|<span data-ttu-id="57544-1180">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1180">Requirement</span></span>|<span data-ttu-id="57544-1181">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1181">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1183">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-1183">1.0</span></span>|
|[<span data-ttu-id="57544-1184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1185">ReadItem</span></span>|
|[<span data-ttu-id="57544-1186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1187">Read</span><span class="sxs-lookup"><span data-stu-id="57544-1187">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57544-1188">Retorna:</span><span class="sxs-lookup"><span data-stu-id="57544-1188">Returns:</span></span>

<span data-ttu-id="57544-p164">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="57544-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="57544-1191">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="57544-1191">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="57544-1192">getInitializationContextAsync ([opções], [callback])</span><span class="sxs-lookup"><span data-stu-id="57544-1192">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="57544-1193">Obtém dados de inicialização passados quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="57544-1193">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="57544-1194">Este método só é compatível com o Outlook 2016 ou posterior no Windows (clique para executar versões posteriores a 16.0.8413.1000) e Outlook na Web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="57544-1194">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1195">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1195">Parameters</span></span>

|<span data-ttu-id="57544-1196">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1196">Name</span></span>|<span data-ttu-id="57544-1197">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1197">Type</span></span>|<span data-ttu-id="57544-1198">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-1198">Attributes</span></span>|<span data-ttu-id="57544-1199">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1199">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="57544-1200">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1200">Object</span></span>|<span data-ttu-id="57544-1201">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1201">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1202">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-1202">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57544-1203">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1203">Object</span></span>|<span data-ttu-id="57544-1204">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1204">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1205">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1205">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57544-1206">function</span><span class="sxs-lookup"><span data-stu-id="57544-1206">function</span></span>|<span data-ttu-id="57544-1207">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1207">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1208">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-1208">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="57544-1209">Com êxito, os dados de inicialização são fornecidos na `asyncResult.value` Propriedade como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="57544-1209">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="57544-1210">Se não houver nenhum contexto de inicialização, `asyncResult` o objeto conterá `Error` um objeto com `code` sua propriedade definida `9020` como e `name` sua propriedade definida `GenericResponseError`como.</span><span class="sxs-lookup"><span data-stu-id="57544-1210">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1211">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1211">Requirements</span></span>

|<span data-ttu-id="57544-1212">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1212">Requirement</span></span>|<span data-ttu-id="57544-1213">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1213">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1214">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1215">Visualização</span><span class="sxs-lookup"><span data-stu-id="57544-1215">Preview</span></span>|
|[<span data-ttu-id="57544-1216">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1217">ReadItem</span></span>|
|[<span data-ttu-id="57544-1218">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1219">Read</span><span class="sxs-lookup"><span data-stu-id="57544-1219">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-1220">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1220">Example</span></span>

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

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="57544-1221">getItemIdAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="57544-1221">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="57544-1222">Obtém de forma assíncrona a ID de um item salvo.</span><span class="sxs-lookup"><span data-stu-id="57544-1222">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="57544-1223">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="57544-1223">Compose mode only.</span></span>

<span data-ttu-id="57544-1224">Quando invocado, este método retorna a ID do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1224">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-1225">Se seu suplemento chamar `getItemIdAsync` um item no modo de redação (por exemplo, para `itemId` usar com o EWS ou a API REST), lembre-se de que, quando o Outlook estiver no modo cache, pode levar algum tempo para que o item seja sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="57544-1225">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="57544-1226">Até que o item seja sincronizado, `itemId` o não é reconhecido e usado retorna um erro.</span><span class="sxs-lookup"><span data-stu-id="57544-1226">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1227">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1227">Parameters</span></span>

|<span data-ttu-id="57544-1228">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1228">Name</span></span>|<span data-ttu-id="57544-1229">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1229">Type</span></span>|<span data-ttu-id="57544-1230">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-1230">Attributes</span></span>|<span data-ttu-id="57544-1231">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1231">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="57544-1232">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1232">Object</span></span>|<span data-ttu-id="57544-1233">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1233">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1234">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-1234">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57544-1235">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1235">Object</span></span>|<span data-ttu-id="57544-1236">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1236">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1237">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1237">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57544-1238">function</span><span class="sxs-lookup"><span data-stu-id="57544-1238">function</span></span>||<span data-ttu-id="57544-1239">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-1239">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="57544-1240">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="57544-1240">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="57544-1241">Erros</span><span class="sxs-lookup"><span data-stu-id="57544-1241">Errors</span></span>

|<span data-ttu-id="57544-1242">Código de erro</span><span class="sxs-lookup"><span data-stu-id="57544-1242">Error code</span></span>|<span data-ttu-id="57544-1243">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1243">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="57544-1244">A ID não pode ser recuperada até que o item seja salvo.</span><span class="sxs-lookup"><span data-stu-id="57544-1244">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1245">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1245">Requirements</span></span>

|<span data-ttu-id="57544-1246">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1246">Requirement</span></span>|<span data-ttu-id="57544-1247">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1247">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1248">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1248">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1249">Visualização</span><span class="sxs-lookup"><span data-stu-id="57544-1249">Preview</span></span>|
|[<span data-ttu-id="57544-1250">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1250">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1251">ReadItem</span></span>|
|[<span data-ttu-id="57544-1252">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1252">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1253">Escrever</span><span class="sxs-lookup"><span data-stu-id="57544-1253">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="57544-1254">Exemplos</span><span class="sxs-lookup"><span data-stu-id="57544-1254">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="57544-1255">O exemplo a seguir mostra a estrutura do `result` parâmetro que é passado para a função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1255">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="57544-1256">A `value` propriedade contém a ID do item.</span><span class="sxs-lookup"><span data-stu-id="57544-1256">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="57544-1257">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="57544-1257">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="57544-1258">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="57544-1258">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-1259">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="57544-1259">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="57544-p168">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="57544-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="57544-1263">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="57544-1263">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="57544-1264">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="57544-1264">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="57544-p169">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="57544-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-1268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1268">Requirements</span></span>

|<span data-ttu-id="57544-1269">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1269">Requirement</span></span>|<span data-ttu-id="57544-1270">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1271">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-1272">1.0</span></span>|
|[<span data-ttu-id="57544-1273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1274">ReadItem</span></span>|
|[<span data-ttu-id="57544-1275">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1276">Read</span><span class="sxs-lookup"><span data-stu-id="57544-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57544-1277">Retorna:</span><span class="sxs-lookup"><span data-stu-id="57544-1277">Returns:</span></span>

<span data-ttu-id="57544-p170">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="57544-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="57544-1280">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="57544-1280">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="57544-1281">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1281">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="57544-1282">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1282">Example</span></span>

<span data-ttu-id="57544-1283">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="57544-1283">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="57544-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="57544-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="57544-1285">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="57544-1285">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-1286">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="57544-1286">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="57544-1287">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="57544-1287">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="57544-p171">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="57544-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1290">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1290">Parameters</span></span>

|<span data-ttu-id="57544-1291">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1291">Name</span></span>|<span data-ttu-id="57544-1292">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1292">Type</span></span>|<span data-ttu-id="57544-1293">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1293">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="57544-1294">String</span><span class="sxs-lookup"><span data-stu-id="57544-1294">String</span></span>|<span data-ttu-id="57544-1295">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="57544-1295">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1296">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1296">Requirements</span></span>

|<span data-ttu-id="57544-1297">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1297">Requirement</span></span>|<span data-ttu-id="57544-1298">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1298">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1299">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1299">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1300">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-1300">1.0</span></span>|
|[<span data-ttu-id="57544-1301">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1301">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1302">ReadItem</span></span>|
|[<span data-ttu-id="57544-1303">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1303">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1304">Read</span><span class="sxs-lookup"><span data-stu-id="57544-1304">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57544-1305">Retorna:</span><span class="sxs-lookup"><span data-stu-id="57544-1305">Returns:</span></span>

<span data-ttu-id="57544-1306">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="57544-1306">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="57544-1307">Tipo: cadeia de caracteres de matriz. < ></span><span class="sxs-lookup"><span data-stu-id="57544-1307">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="57544-1308">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1308">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="57544-1309">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="57544-1309">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="57544-1310">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="57544-1310">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="57544-p172">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="57544-p172">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1313">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1313">Parameters</span></span>

|<span data-ttu-id="57544-1314">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1314">Name</span></span>|<span data-ttu-id="57544-1315">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1315">Type</span></span>|<span data-ttu-id="57544-1316">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-1316">Attributes</span></span>|<span data-ttu-id="57544-1317">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1317">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="57544-1318">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="57544-1318">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="57544-p173">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="57544-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="57544-1322">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1322">Object</span></span>|<span data-ttu-id="57544-1323">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1323">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1324">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-1324">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57544-1325">Object</span><span class="sxs-lookup"><span data-stu-id="57544-1325">Object</span></span>|<span data-ttu-id="57544-1326">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1326">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1327">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1327">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57544-1328">function</span><span class="sxs-lookup"><span data-stu-id="57544-1328">function</span></span>||<span data-ttu-id="57544-1329">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-1329">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="57544-1330">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="57544-1330">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="57544-1331">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="57544-1331">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1332">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1332">Requirements</span></span>

|<span data-ttu-id="57544-1333">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1333">Requirement</span></span>|<span data-ttu-id="57544-1334">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1334">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1335">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1336">1.2</span><span class="sxs-lookup"><span data-stu-id="57544-1336">1.2</span></span>|
|[<span data-ttu-id="57544-1337">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1338">ReadItem</span></span>|
|[<span data-ttu-id="57544-1339">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1340">Escrever</span><span class="sxs-lookup"><span data-stu-id="57544-1340">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="57544-1341">Retorna:</span><span class="sxs-lookup"><span data-stu-id="57544-1341">Returns:</span></span>

<span data-ttu-id="57544-1342">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="57544-1342">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="57544-1343">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="57544-1343">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="57544-1344">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1344">Example</span></span>

```js
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

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="57544-1345">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="57544-1345">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="57544-1346">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="57544-1346">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="57544-1347">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="57544-1347">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="57544-1348">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="57544-1348">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-1349">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1349">Requirements</span></span>

|<span data-ttu-id="57544-1350">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1350">Requirement</span></span>|<span data-ttu-id="57544-1351">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1351">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1352">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1352">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1353">1.6</span><span class="sxs-lookup"><span data-stu-id="57544-1353">1.6</span></span>|
|[<span data-ttu-id="57544-1354">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1354">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1355">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1355">ReadItem</span></span>|
|[<span data-ttu-id="57544-1356">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1356">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1357">Read</span><span class="sxs-lookup"><span data-stu-id="57544-1357">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57544-1358">Retorna:</span><span class="sxs-lookup"><span data-stu-id="57544-1358">Returns:</span></span>

<span data-ttu-id="57544-1359">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="57544-1359">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="57544-1360">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1360">Example</span></span>

<span data-ttu-id="57544-1361">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="57544-1361">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="57544-1362">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="57544-1362">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="57544-p176">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="57544-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="57544-1365">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="57544-1365">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="57544-p177">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="57544-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="57544-1369">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="57544-1369">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="57544-1370">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="57544-1370">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="57544-p178">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="57544-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="57544-1374">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1374">Requirements</span></span>

|<span data-ttu-id="57544-1375">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1375">Requirement</span></span>|<span data-ttu-id="57544-1376">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1376">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1377">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1378">1.6</span><span class="sxs-lookup"><span data-stu-id="57544-1378">1.6</span></span>|
|[<span data-ttu-id="57544-1379">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1380">ReadItem</span></span>|
|[<span data-ttu-id="57544-1381">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1382">Read</span><span class="sxs-lookup"><span data-stu-id="57544-1382">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="57544-1383">Retorna:</span><span class="sxs-lookup"><span data-stu-id="57544-1383">Returns:</span></span>

<span data-ttu-id="57544-p179">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="57544-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="57544-1386">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1386">Example</span></span>

<span data-ttu-id="57544-1387">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="57544-1387">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="57544-1388">getSharedPropertiesAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="57544-1388">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="57544-1389">Obtém as propriedades do compromisso ou da mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="57544-1389">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1390">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1390">Parameters</span></span>

|<span data-ttu-id="57544-1391">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1391">Name</span></span>|<span data-ttu-id="57544-1392">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1392">Type</span></span>|<span data-ttu-id="57544-1393">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-1393">Attributes</span></span>|<span data-ttu-id="57544-1394">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1394">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="57544-1395">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1395">Object</span></span>|<span data-ttu-id="57544-1396">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1396">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1397">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-1397">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57544-1398">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1398">Object</span></span>|<span data-ttu-id="57544-1399">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1399">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1400">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1400">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57544-1401">function</span><span class="sxs-lookup"><span data-stu-id="57544-1401">function</span></span>||<span data-ttu-id="57544-1402">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-1402">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="57544-1403">As propriedades compartilhadas são fornecidas [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) como um objeto `asyncResult.value` na propriedade.</span><span class="sxs-lookup"><span data-stu-id="57544-1403">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="57544-1404">Este objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="57544-1404">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1405">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1405">Requirements</span></span>

|<span data-ttu-id="57544-1406">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1406">Requirement</span></span>|<span data-ttu-id="57544-1407">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1407">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1408">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1409">Visualização</span><span class="sxs-lookup"><span data-stu-id="57544-1409">Preview</span></span>|
|[<span data-ttu-id="57544-1410">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1410">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1411">ReadItem</span></span>|
|[<span data-ttu-id="57544-1412">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-1412">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1413">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-1413">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-1414">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1414">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="57544-1415">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="57544-1415">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="57544-1416">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="57544-1416">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="57544-p181">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="57544-p181">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1420">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1420">Parameters</span></span>

|<span data-ttu-id="57544-1421">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1421">Name</span></span>|<span data-ttu-id="57544-1422">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1422">Type</span></span>|<span data-ttu-id="57544-1423">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-1423">Attributes</span></span>|<span data-ttu-id="57544-1424">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1424">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="57544-1425">function</span><span class="sxs-lookup"><span data-stu-id="57544-1425">function</span></span>||<span data-ttu-id="57544-1426">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-1426">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="57544-1427">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="57544-1427">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="57544-1428">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="57544-1428">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="57544-1429">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1429">Object</span></span>|<span data-ttu-id="57544-1430">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1430">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1431">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1431">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="57544-1432">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1432">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1433">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1433">Requirements</span></span>

|<span data-ttu-id="57544-1434">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1434">Requirement</span></span>|<span data-ttu-id="57544-1435">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1435">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1436">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1436">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1437">1.0</span><span class="sxs-lookup"><span data-stu-id="57544-1437">1.0</span></span>|
|[<span data-ttu-id="57544-1438">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1438">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1439">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1439">ReadItem</span></span>|
|[<span data-ttu-id="57544-1440">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-1440">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1441">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-1441">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-1442">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1442">Example</span></span>

<span data-ttu-id="57544-p184">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="57544-p184">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="57544-1446">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="57544-1446">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="57544-1447">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="57544-1447">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="57544-1448">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="57544-1448">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="57544-1449">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="57544-1449">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="57544-1450">No Outlook na Web e dispositivos móveis, o identificador de anexo é válido somente dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="57544-1450">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="57544-1451">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="57544-1451">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1452">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1452">Parameters</span></span>

|<span data-ttu-id="57544-1453">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1453">Name</span></span>|<span data-ttu-id="57544-1454">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1454">Type</span></span>|<span data-ttu-id="57544-1455">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-1455">Attributes</span></span>|<span data-ttu-id="57544-1456">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1456">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="57544-1457">String</span><span class="sxs-lookup"><span data-stu-id="57544-1457">String</span></span>||<span data-ttu-id="57544-1458">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="57544-1458">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="57544-1459">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1459">Object</span></span>|<span data-ttu-id="57544-1460">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1460">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1461">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-1461">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57544-1462">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1462">Object</span></span>|<span data-ttu-id="57544-1463">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1463">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1464">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1464">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57544-1465">function</span><span class="sxs-lookup"><span data-stu-id="57544-1465">function</span></span>|<span data-ttu-id="57544-1466">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1466">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1467">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-1467">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="57544-1468">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="57544-1468">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="57544-1469">Erros</span><span class="sxs-lookup"><span data-stu-id="57544-1469">Errors</span></span>

|<span data-ttu-id="57544-1470">Código de erro</span><span class="sxs-lookup"><span data-stu-id="57544-1470">Error code</span></span>|<span data-ttu-id="57544-1471">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1471">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="57544-1472">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="57544-1472">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1473">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1473">Requirements</span></span>

|<span data-ttu-id="57544-1474">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1474">Requirement</span></span>|<span data-ttu-id="57544-1475">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1475">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1476">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1477">1.1</span><span class="sxs-lookup"><span data-stu-id="57544-1477">1.1</span></span>|
|[<span data-ttu-id="57544-1478">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1479">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57544-1479">ReadWriteItem</span></span>|
|[<span data-ttu-id="57544-1480">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1481">Escrever</span><span class="sxs-lookup"><span data-stu-id="57544-1481">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-1482">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1482">Example</span></span>

<span data-ttu-id="57544-1483">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="57544-1483">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="57544-1484">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="57544-1484">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="57544-1485">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="57544-1485">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="57544-1486">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="57544-1486">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1487">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1487">Parameters</span></span>

| <span data-ttu-id="57544-1488">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1488">Name</span></span> | <span data-ttu-id="57544-1489">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1489">Type</span></span> | <span data-ttu-id="57544-1490">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-1490">Attributes</span></span> | <span data-ttu-id="57544-1491">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1491">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="57544-1492">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="57544-1492">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="57544-1493">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="57544-1493">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="57544-1494">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1494">Object</span></span> | <span data-ttu-id="57544-1495">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1495">&lt;optional&gt;</span></span> | <span data-ttu-id="57544-1496">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-1496">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="57544-1497">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1497">Object</span></span> | <span data-ttu-id="57544-1498">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1498">&lt;optional&gt;</span></span> | <span data-ttu-id="57544-1499">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1499">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="57544-1500">function</span><span class="sxs-lookup"><span data-stu-id="57544-1500">function</span></span>| <span data-ttu-id="57544-1501">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1501">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1502">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-1502">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1503">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1503">Requirements</span></span>

|<span data-ttu-id="57544-1504">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1504">Requirement</span></span>| <span data-ttu-id="57544-1505">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1505">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1506">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57544-1507">1.7</span><span class="sxs-lookup"><span data-stu-id="57544-1507">1.7</span></span> |
|[<span data-ttu-id="57544-1508">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1508">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="57544-1509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="57544-1509">ReadItem</span></span> |
|[<span data-ttu-id="57544-1510">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="57544-1510">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57544-1511">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="57544-1511">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="57544-1512">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="57544-1512">saveAsync([options], callback)</span></span>

<span data-ttu-id="57544-1513">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="57544-1513">Asynchronously saves an item.</span></span>

<span data-ttu-id="57544-1514">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1514">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="57544-1515">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="57544-1515">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="57544-1516">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="57544-1516">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-1517">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="57544-1517">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="57544-1518">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="57544-1518">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="57544-p188">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="57544-p188">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="57544-1522">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="57544-1522">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="57544-1523">O Outlook no Mac não dá suporte à gravação de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="57544-1523">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="57544-1524">O `saveAsync` método falha quando chamado de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="57544-1524">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="57544-1525">Consulte [não é possível salvar uma reunião como rascunho no Outlook para Mac usando a API do Office js](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="57544-1525">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="57544-1526">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="57544-1526">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1527">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1527">Parameters</span></span>

|<span data-ttu-id="57544-1528">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1528">Name</span></span>|<span data-ttu-id="57544-1529">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1529">Type</span></span>|<span data-ttu-id="57544-1530">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-1530">Attributes</span></span>|<span data-ttu-id="57544-1531">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1531">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="57544-1532">Object</span><span class="sxs-lookup"><span data-stu-id="57544-1532">Object</span></span>|<span data-ttu-id="57544-1533">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1533">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1534">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-1534">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57544-1535">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1535">Object</span></span>|<span data-ttu-id="57544-1536">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1536">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1537">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1537">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="57544-1538">function</span><span class="sxs-lookup"><span data-stu-id="57544-1538">function</span></span>||<span data-ttu-id="57544-1539">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-1539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="57544-1540">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="57544-1540">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1541">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1541">Requirements</span></span>

|<span data-ttu-id="57544-1542">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1542">Requirement</span></span>|<span data-ttu-id="57544-1543">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1543">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1544">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1544">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1545">1.3</span><span class="sxs-lookup"><span data-stu-id="57544-1545">1.3</span></span>|
|[<span data-ttu-id="57544-1546">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1546">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1547">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57544-1547">ReadWriteItem</span></span>|
|[<span data-ttu-id="57544-1548">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1548">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1549">Escrever</span><span class="sxs-lookup"><span data-stu-id="57544-1549">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="57544-1550">Exemplos</span><span class="sxs-lookup"><span data-stu-id="57544-1550">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="57544-p190">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="57544-p190">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="57544-1553">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="57544-1553">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="57544-1554">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="57544-1554">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="57544-p191">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="57544-p191">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="57544-1558">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="57544-1558">Parameters</span></span>

|<span data-ttu-id="57544-1559">Nome</span><span class="sxs-lookup"><span data-stu-id="57544-1559">Name</span></span>|<span data-ttu-id="57544-1560">Tipo</span><span class="sxs-lookup"><span data-stu-id="57544-1560">Type</span></span>|<span data-ttu-id="57544-1561">Atributos</span><span class="sxs-lookup"><span data-stu-id="57544-1561">Attributes</span></span>|<span data-ttu-id="57544-1562">Descrição</span><span class="sxs-lookup"><span data-stu-id="57544-1562">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="57544-1563">String</span><span class="sxs-lookup"><span data-stu-id="57544-1563">String</span></span>||<span data-ttu-id="57544-p192">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="57544-p192">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="57544-1567">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1567">Object</span></span>|<span data-ttu-id="57544-1568">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1568">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1569">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="57544-1569">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="57544-1570">Objeto</span><span class="sxs-lookup"><span data-stu-id="57544-1570">Object</span></span>|<span data-ttu-id="57544-1571">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1571">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1572">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="57544-1572">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="57544-1573">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="57544-1573">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="57544-1574">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="57544-1574">&lt;optional&gt;</span></span>|<span data-ttu-id="57544-1575">Se `text`, o estilo atual é aplicado no Outlook na Web e clientes da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="57544-1575">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="57544-1576">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="57544-1576">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="57544-1577">Se `html` e o campo oferecer suporte a HTML (o assunto não), o estilo atual será aplicado no Outlook na Web e o estilo padrão será aplicado nos clientes da área de trabalho do Outlook.</span><span class="sxs-lookup"><span data-stu-id="57544-1577">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="57544-1578">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="57544-1578">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="57544-1579">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="57544-1579">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="57544-1580">function</span><span class="sxs-lookup"><span data-stu-id="57544-1580">function</span></span>||<span data-ttu-id="57544-1581">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="57544-1581">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57544-1582">Requisitos</span><span class="sxs-lookup"><span data-stu-id="57544-1582">Requirements</span></span>

|<span data-ttu-id="57544-1583">Requisito</span><span class="sxs-lookup"><span data-stu-id="57544-1583">Requirement</span></span>|<span data-ttu-id="57544-1584">Valor</span><span class="sxs-lookup"><span data-stu-id="57544-1584">Value</span></span>|
|---|---|
|[<span data-ttu-id="57544-1585">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="57544-1585">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="57544-1586">1.2</span><span class="sxs-lookup"><span data-stu-id="57544-1586">1.2</span></span>|
|[<span data-ttu-id="57544-1587">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="57544-1587">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="57544-1588">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="57544-1588">ReadWriteItem</span></span>|
|[<span data-ttu-id="57544-1589">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="57544-1589">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="57544-1590">Escrever</span><span class="sxs-lookup"><span data-stu-id="57544-1590">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="57544-1591">Exemplo</span><span class="sxs-lookup"><span data-stu-id="57544-1591">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
